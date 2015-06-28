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

namespace NPOI.POIFS.FileSystem
{
    using System;
    using System.IO;
    using NPOI.Util;
    using System.Text;

    /**
     * Represents an Ole10Native record which is wrapped around certain binary
     * files being embedded in OLE2 documents.
     *
     * @author Rainer Schwarze
     */
    public class Ole10Native
    {
        public static String OLE10_NATIVE = "\x0001Ole10Native";
        protected static String ISO1 = "ISO-8859-1";

        // (the fields as they appear in the raw record:)
        private int totalSize;             // 4 bytes, total size of record not including this field
        private short flags1 = 2;          // 2 bytes, unknown, mostly [02 00]
        private String label;              // ASCIIZ, stored in this field without the terminating zero
        private String fileName;           // ASCIIZ, stored in this field without the terminating zero
        private short flags2 = 0;          // 2 bytes, unknown, mostly [00 00]
        private short unknown1 = 3;        // see below
        private String command;            // ASCIIZ, stored in this field without the terminating zero
        private byte[] dataBuffer;         // varying size, the actual native data
        private short flags3 = 0;          // some final flags? or zero terminators?, sometimes not there


        /// <summary>
        /// Creates an instance of this class from an embedded OLE Object. The OLE Object is expected
        /// to include a stream &quot;{01}Ole10Native&quot; which Contains the actual
        /// data relevant for this class.
        /// </summary>
        /// <param name="poifs">poifs POI Filesystem object</param>
        /// <returns>Returns an instance of this class</returns>
        public static Ole10Native CreateFromEmbeddedOleObject(POIFSFileSystem poifs)
        {
            return CreateFromEmbeddedOleObject(poifs.Root);
        }

        /// <summary>
        /// Creates an instance of this class from an embedded OLE Object. The OLE Object is expected
        /// to include a stream &quot;{01}Ole10Native&quot; which contains the actual
        /// data relevant for this class.
        /// </summary>
        /// <param name="directory">directory POI Filesystem object</param>
        /// <returns>Returns an instance of this class</returns>
        public static Ole10Native CreateFromEmbeddedOleObject(DirectoryNode directory)
        {
            bool plain = false;

            try
            {
                directory.GetEntry("\u0001Ole10ItemName");
                plain = true;
            }
            catch (FileNotFoundException)
            {
                plain = false;
            }

            DocumentEntry nativeEntry =
               (DocumentEntry)directory.GetEntry(OLE10_NATIVE);
            byte[] data = new byte[nativeEntry.Size];
            directory.CreateDocumentInputStream(nativeEntry).Read(data);

            return new Ole10Native(data, 0, plain);
        }
        /**
       * Creates an instance and fills the fields based on ... the fields
       */
        public Ole10Native(String label, String filename, String command, byte[] data)
        {
            Label=(label);
            FileName=(filename);
            Command=(command);
            DataBuffer=(data);
        }

        /**
         * Creates an instance and Fills the fields based on the data in the given buffer.
         *
         * @param data   The buffer Containing the Ole10Native record
         * @param offset The start offset of the record in the buffer
         * @throws Ole10NativeException on invalid or unexcepted data format
         */
        public Ole10Native(byte[] data, int offset)
            : this(data, offset, false)
        {

        }
        /**
         * Creates an instance and Fills the fields based on the data in the given buffer.
         *
         * @param data   The buffer Containing the Ole10Native record
         * @param offset The start offset of the record in the buffer
         * @param plain Specified 'plain' format without filename
         * @throws Ole10NativeException on invalid or unexcepted data format
         */
        public Ole10Native(byte[] data, int offset, bool plain)
        {
            int ofs = offset;        // current offset, Initialized to start

            if (data.Length < offset + 2)
            {
                throw new Ole10NativeException("data is too small");
            }

            totalSize = LittleEndian.GetInt(data, ofs);
            ofs += LittleEndianConsts.INT_SIZE;

            if (plain)
            {
                dataBuffer = new byte[totalSize - 4];
                Array.Copy(data, 4, dataBuffer, 0, dataBuffer.Length);
                //dataSize = totalSize - 4;

                byte[] oleLabel = new byte[8];
                Array.Copy(dataBuffer, 0, oleLabel, 0, Math.Min(dataBuffer.Length, 8));
                label = "ole-" + HexDump.ToHex(oleLabel);
                fileName = label;
                command = label;
            }
            else
            {
                flags1 = LittleEndian.GetShort(data, ofs);
                ofs += LittleEndianConsts.SHORT_SIZE;
                int len = GetStringLength(data, ofs);
                label = StringUtil.GetFromCompressedUnicode(data, ofs, len - 1);
                ofs += len;
                len = GetStringLength(data, ofs);
                fileName = StringUtil.GetFromCompressedUnicode(data, ofs, len - 1);
                ofs += len;
                flags2 = LittleEndian.GetShort(data, ofs);
                ofs += LittleEndianConsts.SHORT_SIZE;

                unknown1 = LittleEndian.GetShort(data, ofs);
                ofs += LittleEndianConsts.SHORT_SIZE;

                len = LittleEndian.GetInt(data, ofs);
                ofs += LittleEndianConsts.INT_SIZE;

                command = StringUtil.GetFromCompressedUnicode(data, ofs, len - 1);
                ofs += len;
                if (totalSize < ofs)
                {
                    throw new Ole10NativeException("Invalid Ole10Native");
                }

                int dataSize = LittleEndian.GetInt(data, ofs);
                ofs += LittleEndianConsts.INT_SIZE;

                if (dataSize < 0 || totalSize - (ofs - LittleEndianConsts.INT_SIZE) < dataSize)
                {
                    throw new Ole10NativeException("Invalid Ole10Native");
                }

                dataBuffer = new byte[dataSize];
                Array.Copy(data, ofs, dataBuffer, 0, dataSize);
                ofs += dataSize;
            }
        }

        /*
         * Helper - determine length of zero terminated string (ASCIIZ).
         */
        private static int GetStringLength(byte[] data, int ofs)
        {
            int len = 0;
            while (len + ofs < data.Length && data[ofs + len] != 0)
            {
                len++;
            }
            len++;
            return len;
        }

        /**
         * Returns the value of the totalSize field - the total length of the structure
         * is totalSize + 4 (value of this field + size of this field).
         *
         * @return the totalSize
         */
        public int TotalSize
        {
            get
            {
                return totalSize;
            }
        }

        /**
         * Returns flags1 - currently unknown - usually 0x0002.
         *
         * @return the flags1
         */
        public short Flags1
        {
            get
            {
                return flags1;
            }
            set { flags1 = value; }
        }

        /**
         * Returns the label field - usually the name of the file (without directory) but
         * probably may be any name specified during packaging/embedding the data.
         *
         * @return the label
         */
        public String Label
        {
            get
            {
                return label;
            }
            set
            {
                label = value;
            }
        }

        /**
         * Returns the fileName field - usually the name of the file being embedded
         * including the full path.
         *
         * @return the fileName
         */
        public String FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
            }
        }

        /**
         * Returns flags2 - currently unknown - mostly 0x0000.
         *
         * @return the flags2
         */
        public short Flags2
        {
            get
            {
                return flags2;
            }
            set
            {
                flags2 = value;
            }
        }

        /**
         * Returns unknown1 field - currently unknown.
         *
         * @return the unknown1
         */
        public short Unknown1
        {
            get
            {
                return unknown1;
            }
            set
            {
                unknown1 = value;
            }
        }

        /**
         * Returns the command field - usually the name of the file being embedded
         * including the full path, may be a command specified during embedding the file.
         *
         * @return the command
         */
        public String Command
        {
            get
            {
                return command;
            }
            set { command = value; }
        }

        /**
         * Returns the size of the embedded file. If the size is 0 (zero), no data has been
         * embedded. To be sure, that no data has been embedded, check whether
         * {@link #getDataBuffer()} returns <code>null</code>.
         *
         * @return the dataSize
         */
        public int DataSize
        {
            get{return dataBuffer.Length;}
            
        }

        /**
         * Returns the buffer Containing the embedded file's data, or <code>null</code>
         * if no data was embedded. Note that an embedding may provide information about
         * the data, but the actual data is not included. (So label, filename etc. are
         * available, but this method returns <code>null</code>.)
         *
         * @return the dataBuffer
         */
        public byte[] DataBuffer
        {
            get{return dataBuffer;}
            set { dataBuffer = value; }
        }

        /**
         * Returns the flags3 - currently unknown.
         *
         * @return the flags3
         */
        public short Flags3
        {
            get
            {
                return flags3;
            }
            set
            {
                flags3 = value;
            }
        }

        /**
   * Have the contents printer out into an OutputStream, used when writing a
   * file back out to disk (Normally, atom classes will keep their bytes
   * around, but non atom classes will just request the bytes from their
   * children, then chuck on their header and return)
   */
        public void WriteOut(Stream out1)
        {
            byte[] intbuf = new byte[LittleEndianConsts.INT_SIZE];
            byte[] shortbuf = new byte[LittleEndianConsts.SHORT_SIZE];
            byte[] zerobuf = { 0, 0, 0, 0 };
            MemoryStream bos = new MemoryStream();
            bos.Write(intbuf, 0, intbuf.Length); // total size, will be determined later ..

            LittleEndian.PutShort(shortbuf, 0, Flags1);
            bos.Write(shortbuf, 0, shortbuf.Length);
            byte[] buffer = Encoding.GetEncoding(ISO1).GetBytes(Label);
            bos.Write(buffer, 0, buffer.Length);
            bos.Write(zerobuf, 0, zerobuf.Length);
            buffer = Encoding.GetEncoding(ISO1).GetBytes(FileName);
            bos.Write(buffer, 0, buffer.Length);
            bos.Write(zerobuf, 0, zerobuf.Length);

            LittleEndian.PutShort(shortbuf, 0, Flags2);
            bos.Write(shortbuf, 0, shortbuf.Length);

            LittleEndian.PutShort(shortbuf, 0, Unknown1);
            bos.Write(shortbuf, 0, shortbuf.Length);

            LittleEndian.PutInt(intbuf, 0, Command.Length+1);
            bos.Write(intbuf, 0, intbuf.Length);
            buffer = Encoding.GetEncoding(ISO1).GetBytes(Command);
            bos.Write(buffer, 0, buffer.Length);
            bos.Write(zerobuf, 0, zerobuf.Length);

            LittleEndian.PutInt(intbuf, 0, DataBuffer.Length);
            bos.Write(intbuf, 0, intbuf.Length);

            bos.Write(DataBuffer, 0, DataBuffer.Length);

            LittleEndian.PutShort(shortbuf, 0, Flags3);
            bos.Write(shortbuf, 0, shortbuf.Length);

            // update total size - length of length-field (4 bytes)
            byte[] data = bos.ToArray();
            totalSize = data.Length - LittleEndianConsts.INT_SIZE;
            LittleEndian.PutInt(data, 0, totalSize);

            out1.Write(data, 0, data.Length);
        }

    }

}





