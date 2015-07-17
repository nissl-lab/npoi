
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
     * Title: Codepage Record
     * <p>Description:  the default characterset. for the workbook</p>
     * <p>REFERENCE:  PG 293 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)</p>
     * <p>Use {@link CodePageUtil} to turn these values into Java code pages
     *  to encode/decode strings.</p>
     * @version 2.0-pre
     */
    public class CodepageRecord
       : StandardRecord
    {
        public const short sid = 0x42;
        private short field_1_codepage;   // = 0;

        /**
         * Excel 97+ (Biff 8) should always store strings as UTF-16LE or
         *  compressed versions of that. As such, this should always be
         *  0x4b0 = UTF_16, except for files coming from older versions.
         */
        public const short CODEPAGE = (short)0x4b0;

        public CodepageRecord()
        {
        }

        /**
         * Constructs a CodepageRecord and Sets its fields appropriately
         * @param in the RecordInputstream to Read the record from
         */

        public CodepageRecord(RecordInputStream in1)
        {
            field_1_codepage = in1.ReadShort();
        }
        /**
         * Get the codepage for this workbook
         *
         * @see #CODEPAGE
         * @return codepage - the codepage to Set
         */

        public short Codepage
        {
            get
            {
                return field_1_codepage;
            }
            set { field_1_codepage = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[CODEPAGE]\n");
            buffer.Append("    .codepage        = ")
                .Append(StringUtil.ToHexString(Codepage)).Append("\n");
            buffer.Append("[/CODEPAGE]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Codepage);
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