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

using System.IO;
using System.Security;
using NPOI.Util;

namespace NPOI.POIFS.Crypt
{
    using System;
    using System.Collections;

    public abstract class ChunkedCipherInputStream : LittleEndianInputStream
    {
        private int chunkSize;
        private int chunkMask;
        private int chunkBits;

        private int _lastIndex = 0;
        private long _pos = 0;
        private long _size;
        private byte[] _chunk;
        private Cipher _cipher;
        protected IEncryptionInfoBuilder builder;
        protected Decryptor decryptor;
        public ChunkedCipherInputStream(ILittleEndianInput stream, long size, int chunkSize
            , IEncryptionInfoBuilder builder, Decryptor decryptor)
            : base((Stream)stream)
        {
            
            _size = size;
            this.chunkSize = chunkSize;
            chunkMask = chunkSize - 1;
            chunkBits = Number.BitCount(chunkMask);
            this.builder = builder;
            this.decryptor = decryptor;
            _cipher = InitCipherForBlock(null, 0);
        }

        protected abstract Cipher InitCipherForBlock(Cipher existing, int block);

        public int Read()
        {
            byte[] b = new byte[1];
            if (Read(b) == 1)
                return b[0];
            return -1;
        }

        // do not implement! -> recursion
        // public int Read(byte[] b) throws IOException;

        public new int Read(byte[] b, int off, int len)
        {
            int total = 0;

            if (Available() <= 0) return -1;

            while (len > 0)
            {
                if (_chunk == null)
                {
                    try
                    {
                        _chunk = NextChunk();
                    }
                    catch (SecurityException e)
                    {
                        throw new EncryptedDocumentException(e.Message, e);
                    }
                }
                int count = (int) (chunkSize - (_pos & chunkMask));
                int avail = Available();
                if (avail == 0)
                {
                    return total;
                }
                count = Math.Min(avail, Math.Min(count, len));
                Array.Copy(_chunk, (int) (_pos & chunkMask), b, off, count);
                off += count;
                len -= count;
                _pos += count;
                if ((_pos & chunkMask) == 0)
                    _chunk = null;
                total += count;
            }

            return total;
        }


        public new long Skip(long n)
        {
            long start = _pos;
            long skip = Math.Min(Available(), n);

            if ((((_pos + skip) ^ start) & ~chunkMask) != 0)
                _chunk = null;
            _pos += skip;
            return skip;
        }


        public new int Available()
        {
            return (int) (_size - _pos);
        }


        public bool MarkSupported()
        {
            return false;
        }


        public void Mark(int Readlimit)
        {
            throw new InvalidOperationException();
        }


        public void Reset()
        {
            throw new InvalidOperationException();
        }

        private byte[] NextChunk()
        {
            int index = (int) (_pos >> chunkBits);
            InitCipherForBlock(_cipher, index);

            if (_lastIndex != index)
            {
                base.Skip((index - _lastIndex) << chunkBits);
            }

            byte[] block = new byte[Math.Min(base.Available(), chunkSize)];
            base.Read(block, 0, block.Length);
            _lastIndex = index + 1;
            return _cipher.DoFinal(block);
        }
    }

}