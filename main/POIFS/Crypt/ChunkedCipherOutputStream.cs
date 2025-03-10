/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
namespace NPOI.POIFS.Crypt
{
    using System;
    using System.IO;
    using NPOI.POIFS.EventFileSystem;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;


    public abstract class ChunkedCipherOutputStream : LittleEndianOutputStream {
        protected int chunkSize;
        protected int chunkMask;
        protected int chunkBits;

        private readonly byte[] _chunk;
        private readonly FileInfo fileOut;
        private readonly DirectoryNode dir;

        private long _pos = 0;
        private Cipher _cipher;
        protected IEncryptionInfoBuilder builder;
        protected Encryptor encryptor;
        public Stream GetStream()
        {
            return out1;
        }
        public ChunkedCipherOutputStream(DirectoryNode dir, int chunkSize, IEncryptionInfoBuilder builder, Encryptor encryptor)
            : base(null)
        {
            this.chunkSize = chunkSize;
            chunkMask = chunkSize - 1;
            chunkBits = Number.BitCount(chunkMask);
            _chunk = new byte[chunkSize];

            fileOut = TempFile.CreateTempFile("encrypted_package", "crypt");
            //fileOut.DeleteOnExit();
            this.out1 = fileOut.Create();
            this.dir = dir;
            this.builder = builder;
            this.encryptor = encryptor;
            _cipher = InitCipherForBlock(null, 0, false);
        }

        protected abstract Cipher InitCipherForBlock(Cipher existing, int block, bool lastChunk);

        protected abstract void CalculateChecksum(FileInfo fileOut, int oleStreamSize);

        protected abstract void CreateEncryptionInfoEntry(DirectoryNode dir, FileInfo tmpFile);

        public void Write(int b) {
            Write(new byte[] { (byte)b });
        }

        public new void Write(byte[] b) {
            Write(b, 0, b.Length);
        }

        public new void Write(byte[] b, int off, int len)
        {
            if (len == 0) return;

            if (len < 0 || b.Length < off + len) {
                throw new IOException("not enough bytes in your input buffer");
            }

            while (len > 0) {
                int posInChunk = (int)(_pos & chunkMask);
                int nextLen = Math.Min(chunkSize - posInChunk, len);
                Array.Copy(b, off, _chunk, posInChunk, nextLen);
                _pos += nextLen;
                off += nextLen;
                len -= nextLen;
                if ((_pos & chunkMask) == 0) {
                    try {
                        WriteChunk();
                    } catch (Exception e) {
                        throw new IOException(e.Message);
                    }
                }
            }
        }

        protected void WriteChunk() {
            int posInChunk = (int)(_pos & chunkMask);
            // normally posInChunk is 0, i.e. on the next chunk (-> index-1)
            // but if called on close(), posInChunk is somewhere within the chunk data
            int index = (int)(_pos >> chunkBits);
            bool lastChunk;
            if (posInChunk == 0) {
                index--;
                posInChunk = chunkSize;
                lastChunk = false;
            } else {
                // pad the last chunk
                lastChunk = true;
            }

            _cipher = InitCipherForBlock(_cipher, index, lastChunk);

            int ciLen = _cipher.DoFinal(_chunk, 0, posInChunk, _chunk);
            out1.Write(_chunk, 0, ciLen);
        }

        public new void Close() {
            try {
                WriteChunk();

                base.Close();

                int oleStreamSize = (int)(fileOut.Length + LittleEndianConsts.LONG_SIZE);
                CalculateChecksum(fileOut, (int)_pos);
                dir.CreateDocument(Decryptor.DEFAULT_POIFS_ENTRY, oleStreamSize, new EncryptedPackageWriter(this));
                CreateEncryptionInfoEntry(dir, fileOut);
            } catch (Exception e) {
                throw new IOException(e.Message);
            }
        }

        private sealed class EncryptedPackageWriter : POIFSWriterListener {
            readonly ChunkedCipherOutputStream stream;
            public EncryptedPackageWriter(ChunkedCipherOutputStream stream)
            {
                this.stream = stream;
            }
            public void ProcessPOIFSWriterEvent(POIFSWriterEvent event1) {
                try {
                    DocumentOutputStream os = event1.Stream;
                    byte[] buf = new byte[stream.chunkSize];

                    // StreamSize (8 bytes): An unsigned integer that specifies the number of bytes used by data 
                    // encrypted within the EncryptedData field, not including the size of the StreamSize field. 
                    // Note that the actual size of the \EncryptedPackage stream (1) can be larger than this 
                    // value, depending on the block size of the chosen encryption algorithm
                    LittleEndian.PutLong(buf, 0, stream._pos);
                    os.Write(buf, 0, LittleEndian.LONG_SIZE);

                    FileStream fis = stream.fileOut.Create();
                    int readBytes;
                    while ((readBytes = fis.Read(buf, 0, buf.Length)) != -1) {
                        os.Write(buf, 0, readBytes);
                    }
                    fis.Close();

                    os.Close();

                    stream.fileOut.Delete();
                } catch (IOException e) {
                    throw new EncryptedDocumentException(e);
                }
            }
        }
    }

}