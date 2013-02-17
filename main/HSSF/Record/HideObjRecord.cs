
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
    using NPOI.Util;
    using System.Text;

    /**
     * Title:        Hide Object Record
     * Description:  flag defines whether to hide placeholders and object
     * REFERENCE:  PG 321 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class HideObjRecord
       : StandardRecord
    {
        public const short sid = 0x8d;
        public const short HIDE_ALL = 2;
        public const short SHOW_PLACEHOLDERS = 1;
        public const short SHOW_ALL = 0;
        private short field_1_hide_obj;

        public HideObjRecord()
        {
        }

        /**
         * Constructs an HideObj record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public HideObjRecord(RecordInputStream in1)
        {
            field_1_hide_obj = in1.ReadShort();
        }


        /**
         * Set hide object options
         *
         * @param hide options
         * @see #HIDE_ALL
         * @see #SHOW_PLACEHOLDERS
         * @see #SHOW_ALL
         */

        public void SetHideObj(short hide)
        {
            field_1_hide_obj = hide;
        }

        /**
         * Get hide object options
         *
         * @return hide options
         * @see #HIDE_ALL
         * @see #SHOW_PLACEHOLDERS
         * @see #SHOW_ALL
         */

        public short GetHideObj()
        {
            return field_1_hide_obj;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[HIDEOBJ]\n");
            buffer.Append("    .hideobj         = ")
                .Append(StringUtil.ToHexString(GetHideObj())).Append("\n");
            buffer.Append("[/HIDEOBJ]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(GetHideObj());
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
