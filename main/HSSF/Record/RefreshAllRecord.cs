
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
     * Title:        Refresh All Record 
     * Description:  Flag whether to refresh all external data when loading a sheet.
     *               (which hssf doesn't support anyhow so who really cares?)
     * REFERENCE:  PG 376 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class RefreshAllRecord
       : Record
    {
        public static short sid = 0x1B7;
        private short field_1_refreshall;

        public RefreshAllRecord()
        {
        }

        /**
         * Constructs a RefreshAll record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public RefreshAllRecord(RecordInputStream in1)
        {
            field_1_refreshall = in1.ReadShort();
        }

        /**
         * Get whether to refresh all external data when loading a sheet
         * @return refreshall or not
         */

        public bool RefreshAll
        {
            get{return (field_1_refreshall == 1);}
            set
            {
                if (value)
                {
                    field_1_refreshall = 1;
                }
                else
                {
                    field_1_refreshall = 0;
                }
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[REFRESHALL]\n");
            buffer.Append("    .refreshall      = ").Append(RefreshAll)
                .Append("\n");
            buffer.Append("[/REFRESHALL]\n");
            return buffer.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutShort(data, 0 + offset, sid);
            LittleEndian.PutShort(data, 2 + offset,
                                  ((short)0x02));   // 2 bytes (6 total)
            LittleEndian.PutShort(data, 4 + offset, field_1_refreshall);
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