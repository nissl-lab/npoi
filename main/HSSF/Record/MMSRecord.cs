
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
    using NPOI.Util;
    using System;
    using System.Text;


    /**
     * Title: MMS Record
     * Description: defines how many Add menu and del menu options are stored
     *                    in the file. Should always be Set to 0 for HSSF workbooks
     * REFERENCE:  PG 328 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class MMSRecord
       : StandardRecord
    {
        public const short sid = 0xC1;
        private byte field_1_AddMenuCount;   // = 0;
        private byte field_2_delMenuCount;   // = 0;

        public MMSRecord()
        {
        }

        /**
         * Constructs a MMS record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public MMSRecord(RecordInputStream in1)
        {
            if (in1.Remaining == 0)
            {
                return;
            }
            field_1_AddMenuCount = (byte)in1.ReadByte();
            field_2_delMenuCount = (byte)in1.ReadByte();
        }

        /**
         * Set number of Add menu options (Set to 0)
         * @param am  number of Add menu options
         */

        public byte AddMenuCount
        {
            get
            {
                return field_1_AddMenuCount;
            }
            set
            {
                field_1_AddMenuCount = value;
            }
        }

        /**
         * Set number of del menu options (Set to 0)
         * @param dm  number of del menu options
         */

        public byte DelMenuCount
        {
            get
            {
                return field_2_delMenuCount;
            }
            set
            {
                field_2_delMenuCount = value;
            }
        }


        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[MMS]\n");
            buffer.Append("    .addMenu        = ")
                .Append(StringUtil.ToHexString(AddMenuCount)).Append("\n");
            buffer.Append("    .delMenu        = ")
                .Append(StringUtil.ToHexString(DelMenuCount)).Append("\n");
            buffer.Append("[/MMS]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(AddMenuCount);
            out1.WriteByte(DelMenuCount);
        }

        protected override int DataSize
        {
            get
            {
                return 2;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}