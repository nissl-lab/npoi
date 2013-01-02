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

    /**
     * Represents an Ole10Native record which is wrapped around certain binary
     * files being embedded in OLE2 documents.
     *
     * @author Rainer Schwarze
     */
    public class Ole10Native
    {
        // (the fields as they appear in the raw record:)
        private int totalSize;                // 4 bytes, total size of record not including this field
        private short flags1;                // 2 bytes, unknown, mostly [02 00]
        private String label;                // ASCIIZ, stored in this field without the terminating zero
        private String fileName;        // ASCIIZ, stored in this field without the terminating zero
        private short flags2;                // 2 bytes, unknown, mostly [00 00]
        // private byte unknown1Length;	// 1 byte, specifying the length of the following byte array (unknown1)
        private byte[] unknown1;        // see below
        private byte[] unknown2;        // 3 bytes, unknown, mostly [00 00 00]
        private String command;                // ASCIIZ, stored in this field without the terminating zero
        private int dataSize;                // 4 bytes (if space), size of following buffer
        private byte[] dataBuffer;        // varying size, the actual native data
        private short flags3;                // some flags? or zero terminators?, sometimes not there
        public static String OLE10_NATIVE = "\x0001Ole10Native";


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
         * Creates an instance and Fills the fields based on the data in the given buffer.
         *
         * @param data   The buffer Containing the Ole10Native record
         * @param offset The start offset of the record in the buffer
         * @throws Ole10NativeException on invalid or unexcepted data format
         */
        public Ole10Native(byte[] data, int offset):this(data, offset, false)
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
                dataSize = totalSize - 4;

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
                len = LittleEndian.GetUByte(data, ofs);
                unknown1 = new byte[len];
                ofs += len;
                len = 3;
                unknown2 = new byte[len];
                ofs += len;
                len = GetStringLength(data, ofs);
                command = StringUtil.GetFromCompressedUnicode(data, ofs, len - 1);
                ofs += len;

                if (totalSize + LittleEndianConsts.INT_SIZE - ofs > LittleEndianConsts.INT_SIZE)
                {
                    dataSize = LittleEndian.GetInt(data, ofs);
                    ofs += LittleEndianConsts.INT_SIZE;

                    if (dataSize > totalSize || dataSize < 0)
                    {
                        throw new Ole10NativeException("Invalid Ole10Native");
                    }

                    dataBuffer = new byte[dataSize];
                    Array.Copy(data, ofs, dataBuffer, 0, dataSize);
                    ofs += dataSize;

                    if (unknown1.Length > 0)
                    {
                        flags3 = LittleEndian.GetShort(data, ofs);
                        ofs += LittleEndianConsts.SHORT_SIZE;
                    }
                    else
                    {
                        flags3 = 0;
                    }
                }
                else
                {
                    throw new Ole10NativeException("Invalid Ole10Native");
                }
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
        public short GetFlags1()
        {
            return flags1;
        }

        /**
         * Returns the label field - usually the name of the file (without directory) but
         * probably may be any name specified during packaging/embedding the data.
         *
         * @return the label
         */
        public String GetLabel()
        {
            return label;
        }

        /**
         * Returns the fileName field - usually the name of the file being embedded
         * including the full path.
         *
         * @return the fileName
         */
        public String GetFileName()
        {
            return fileName;
        }

        /**
         * Returns flags2 - currently unknown - mostly 0x0000.
         *
         * @return the flags2
         */
        public short GetFlags2()
        {
            return flags2;
        }

        /**
         * Returns unknown1 field - currently unknown.
         *
         * @return the unknown1
         */
        public byte[] GetUnknown1()
        {
            return unknown1;
        }

        /**
         * Returns the unknown2 field - currently being a byte[3] - mostly {0, 0, 0}.
         *
         * @return the unknown2
         */
        public byte[] GetUnknown2()
        {
            return unknown2;
        }

        /**
         * Returns the command field - usually the name of the file being embedded
         * including the full path, may be a command specified during embedding the file.
         *
         * @return the command
         */
        public String GetCommand()
        {
            return command;
        }

        /**
         * Returns the size of the embedded file. If the size is 0 (zero), no data has been
         * embedded. To be sure, that no data has been embedded, check whether
         * {@link #getDataBuffer()} returns <code>null</code>.
         *
         * @return the dataSize
         */
        public int GetDataSize()
        {
            return dataSize;
        }

        /**
         * Returns the buffer Containing the embedded file's data, or <code>null</code>
         * if no data was embedded. Note that an embedding may provide information about
         * the data, but the actual data is not included. (So label, filename etc. are
         * available, but this method returns <code>null</code>.)
         *
         * @return the dataBuffer
         */
        public byte[] GetDataBuffer()
        {
            return dataBuffer;
        }

        /**
         * Returns the flags3 - currently unknown.
         *
         * @return the flags3
         */
        public short GetFlags3()
        {
            return flags3;
        }
    }

}





