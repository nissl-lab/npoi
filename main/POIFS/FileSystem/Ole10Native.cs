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

        /**
         * the field encoding mode - merely a try-and-error guess ...
         **/ 
        private enum EncodingMode {
            /**
             * the data is stored in parsed format - including label, command, etc.
             */
            parsed,
            /**
             * the data is stored raw after the length field
             */
            unparsed,
            /**
             * the data is stored raw after the length field and the flags1 field
             */
            compact
        }
    
        private EncodingMode mode;
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
            DocumentEntry nativeEntry =
               (DocumentEntry)directory.GetEntry(OLE10_NATIVE);
            byte[] data = new byte[nativeEntry.Size];
            directory.CreateDocumentInputStream(nativeEntry).Read(data);

            return new Ole10Native(data, 0);
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
            mode = EncodingMode.parsed;
        }

        /**
         * Creates an instance and Fills the fields based on the data in the given buffer.
         *
         * @param data   The buffer Containing the Ole10Native record
         * @param offset The start offset of the record in the buffer
         * @param plain as of POI 3.11 this parameter is ignored
         * @throws Ole10NativeException on invalid or unexcepted data format
         */
        [Obsolete("parameter plain is ignored, use {@link #Ole10Native(byte[],int)}")]
        public Ole10Native(byte[] data, int offset, bool plain)
            : this(data, offset)
        {

        }
        /**
         * Creates an instance and Fills the fields based on the data in the given buffer.
         *
         * @param data   The buffer Containing the Ole10Native record
         * @param offset The start offset of the record in the buffer
         * @throws Ole10NativeException on invalid or unexcepted data format
         */
        public Ole10Native(byte[] data, int offset)
        {
            int ofs = offset;        // current offset, Initialized to start

            if (data.Length < offset + 2)
            {
                throw new Ole10NativeException("data is too small");
            }

            totalSize = LittleEndian.GetInt(data, ofs);
            ofs += LittleEndianConsts.INT_SIZE;
            mode = EncodingMode.unparsed;
            if (LittleEndian.GetShort(data, ofs) == 2)
            {
                // some files like equations don't have a valid filename,
                // but somehow encode the formula right away in the ole10 header
                if (char.IsControl((char)data[ofs + LittleEndianConsts.SHORT_SIZE]))
                {
                    mode = EncodingMode.compact;
                }
                else
                {
                    mode = EncodingMode.parsed;
                }
            }
            int dataSize = 0;
            switch (mode)
            {
                case EncodingMode.parsed:
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

                    dataSize = LittleEndian.GetInt(data, ofs);
                    ofs += LittleEndianConsts.INT_SIZE;

                    if (dataSize < 0 || totalSize - (ofs - LittleEndianConsts.INT_SIZE) < dataSize)
                    {
                        throw new Ole10NativeException("Invalid Ole10Native");
                    }

                    break;
                case EncodingMode.compact:
                    flags1 = LittleEndian.GetShort(data, ofs);
                    ofs += LittleEndianConsts.SHORT_SIZE;
                    dataSize = totalSize - LittleEndianConsts.SHORT_SIZE;
                    break;
                case EncodingMode.unparsed:
                    dataSize = totalSize;
                    break;

            }
            dataBuffer = new byte[dataSize];
            Array.Copy(data, ofs, dataBuffer, 0, dataSize);
            ofs += dataSize;

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
            
            LittleEndianOutputStream leosOut = new LittleEndianOutputStream(out1);

            switch (mode)
            {
                case EncodingMode.parsed:
                    {
                        MemoryStream bos = new MemoryStream();
                        LittleEndianOutputStream leos = new LittleEndianOutputStream(bos);
                        // total size, will be determined later ..

                        leos.WriteShort(Flags1);
                        leos.Write(Encoding.GetEncoding(ISO1).GetBytes(Label));
                        leos.WriteByte(0);
                        leos.Write(Encoding.GetEncoding(ISO1).GetBytes(FileName));
                        leos.WriteByte(0);
                        leos.WriteShort(Flags2);
                        leos.WriteShort(Unknown1);
                        leos.WriteInt(Command.Length + 1);
                        leos.Write(Encoding.GetEncoding(ISO1).GetBytes(Command));
                        leos.WriteByte(0);
                        leos.WriteInt(DataSize);
                        leos.Write(DataBuffer);
                        leos.WriteShort(Flags3);
                        //leos.Close(); // satisfy compiler ...

                        leosOut.WriteInt((int)bos.Length); // total size
                        bos.WriteTo(out1);
                        break;
                    }
                case EncodingMode.compact:
                    leosOut.WriteInt(DataSize + LittleEndianConsts.SHORT_SIZE);
                    leosOut.WriteShort(Flags1);
                    out1.Write(DataBuffer, 0, DataBuffer.Length);
                    break;
                default:
                case EncodingMode.unparsed:
                    leosOut.WriteInt(DataSize);
                    out1.Write(DataBuffer, 0, DataBuffer.Length);
                    break;
            }

        }

    }

}





