/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.Util;

    /**
     * Represents a NoteStructure (0xD) sub record.
     *
     * 
     * The docs say nothing about it. The Length of this record is always 26 bytes.
     * 
     *
     * @author Yegor Kozlov
     */
    public class NoteStructureSubRecord: SubRecord
    {
        public const short sid = 0x0D;
        private const int ENCODED_SIZE = 22;

        private byte[] reserved;

        /**
         * Construct a new <c>NoteStructureSubRecord</c> and
         * Fill its data with the default values
         */
        public NoteStructureSubRecord()
        {
            //all we know is that the the Length of <c>NoteStructureSubRecord</c> is always 22 bytes
            reserved = new byte[ENCODED_SIZE];
        }

        /**
         * Constructs a NoteStructureSubRecord and Sets its fields appropriately.
         *
         */
        public NoteStructureSubRecord(ILittleEndianInput in1, int size)
        {
            if (size != ENCODED_SIZE) {
                throw new RecordFormatException("Unexpected size (" + size + ")");
            }
            //just grab the raw data
            byte[] buf = new byte[size];
            in1.ReadFully(buf);
            reserved = buf;
        }

        /**
         * Convert this record to string.
         * Used by BiffViewer and other utulities.
         */
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            String nl = Environment.NewLine;
            buffer.Append("[ftNts ]" + nl);
            buffer.Append("  size     = ").Append(DataSize).Append(nl);
            buffer.Append("  reserved = ").Append(HexDump.ToHex(reserved)).Append(nl);
            buffer.Append("[/ftNts ]" + nl);
            return buffer.ToString();
        }

        /**
         * Serialize the record data into the supplied array of bytes
         *
         * @param offset offset in the <c>data</c>
         * @param data the data to Serialize into
         *
         * @return size of the record
         */
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(reserved.Length);
            out1.Write(reserved);
        }

        /**
         * Size of record
         */
        public override int DataSize
        {
            get { return reserved.Length; }
        }

        /**
         * @return id of this record.
         */
        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            NoteStructureSubRecord rec = new NoteStructureSubRecord();
            byte[] recdata = new byte[reserved.Length];
            Array.Copy(reserved, 0, recdata, 0, recdata.Length);
            rec.reserved = recdata;
            return rec;
        }

    }

}