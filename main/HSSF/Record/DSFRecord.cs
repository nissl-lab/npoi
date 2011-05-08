
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
     * Title: double Stream Flag Record
     * Description:  tells if this Is a double stream file. (always no for HSSF generated files)
     *               double Stream files contain both BIFF8 and BIFF7 workbooks.
     * REFERENCE:  PG 305 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class DSFRecord
       : Record
    {
        public const short sid = 0x161;
        private short field_1_dsf;

        public DSFRecord()
        {
        }

        /**
         * Constructs a DBCellRecord and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public DSFRecord(RecordInputStream in1)
        {
            field_1_dsf = in1.ReadShort();
        }

        /**
         * Get the DSF flag
         * @return dsfflag (0-off,1-on)
         */

        public short Dsf
        {
            get
            {
                return field_1_dsf;
            }
            set { field_1_dsf = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DSF]\n");
            buffer.Append("    .IsDSF           = ")
                .Append(StringUtil.ToHexString(Dsf)).Append("\n");
            buffer.Append("[/DSF]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset,
                                  ((short)0x02));   // 2 bytes (6 total)
            LittleEndian.PutShort(data, 4 + offset, Dsf);
            return RecordSize;
        }

        public override int RecordSize
        {
            get { return 6; }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}
