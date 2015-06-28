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

namespace NPOI.HSSF.Record
{
    using System;
    using NPOI.Util;
    using System.Text;


    /**
     * The FtCf structure specifies the clipboard format of the picture-type Obj record Containing this FtCf.
     */
    public class FtCfSubRecord : SubRecord
    {
        public const short sid = 0x07;
        public const short length = 0x02;

        /**
         * Specifies the format of the picture is an enhanced metafile.
         */
        public static short METAFILE_BIT = (short)0x0002;

        /**
         * Specifies the format of the picture is a bitmap.
         */
        public static short BITMAP_BIT = (short)0x0009;

        /**
         * Specifies the picture is in an unspecified format that is
         * neither and enhanced metafile nor a bitmap.
         */
        public static short UNSPECIFIED_BIT = unchecked((short)0xFFFF);

        private short flags = 0;

        /**
         * Construct a new <code>FtPioGrbitSubRecord</code> and
         * fill its data with the default values
         */
        public FtCfSubRecord()
        {
        }

        public FtCfSubRecord(ILittleEndianInput in1, int size)
        {
            if (size != length)
            {
                throw new RecordFormatException("Unexpected size (" + size + ")");
            }
            flags = in1.ReadShort();
        }

        /**
         * Convert this record to string.
         * Used by BiffViewer and other utilities.
         */
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[FtCf ]\n");
            buffer.Append("  size     = ").Append(length).Append("\n");
            buffer.Append("  flags    = ").Append(HexDump.ToHex(flags)).Append("\n");
            buffer.Append("[/FtCf ]\n");
            return buffer.ToString();
        }

        /**
         * Serialize the record data into the supplied array of bytes
         *
         * @param out the stream to serialize into
         */
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(length);
            out1.WriteShort(flags);
        }

        public override int DataSize
        {
            get
            {
                return length;
            }
        }

        /**
         * @return id of this record.
         */
        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public override Object Clone()
        {
            FtCfSubRecord rec = new FtCfSubRecord();
            rec.flags = this.flags;
            return rec;
        }

        public short Flags
        {
            get
            {
                return flags;
            }
            set
            {
                this.flags = value;
            }
        }
    }
}