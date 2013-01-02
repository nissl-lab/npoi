
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
     * Title:        Date Window 1904 Flag record 
     * Description:  Flag specifying whether 1904 date windowing Is used.
     *               (tick toc tick toc...BOOM!) 
     * REFERENCE:  PG 280 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class DateWindow1904Record
       : StandardRecord
    {
        public const short sid = 0x22;
        private short field_1_window;

        public DateWindow1904Record()
        {
        }

        /**
         * Constructs a DateWindow1904 record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public DateWindow1904Record(RecordInputStream in1)
        {
            field_1_window = in1.ReadShort();
        }

        /**
         * Gets whether or not to use 1904 date windowing (which means you'll be screwed in 2004)
         * @return window flag - 0/1 (false,true)
         */

        public short Windowing
        {
            get { return field_1_window; }
            set { field_1_window = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[1904]\n");
            buffer.Append("    .is1904          = ")
                .Append(StringUtil.ToHexString(Windowing)).Append("\n");
            buffer.Append("[/1904]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(Windowing);
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