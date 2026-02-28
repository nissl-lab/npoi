
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
     * Title: Window Protect Record
     * Description:  flags whether workbook windows are protected
     * REFERENCE:  PG 424 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class WindowProtectRecord : StandardRecord
    {
        public const short sid = 0x19;
        //private short field_1_protect;
        private static BitField settingsProtectedFlag = BitFieldFactory.GetInstance(0x0001);
        private int _options;
        public WindowProtectRecord(int options)
        {
            _options = options;
        }

        /**
         * Constructs a WindowProtect record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public WindowProtectRecord(RecordInputStream in1)
            : this(in1.ReadUShort())
        {
        }
        public WindowProtectRecord(bool protect)
            :this(0)
        {
            Protect = (protect);
        }
        /**
         * Is this window protected or not
         *
         * @return protected or not
         */

        public bool Protect
        {
            get
            {
                //return (field_1_protect == 1);
                return settingsProtectedFlag.IsSet(_options);
            }
            set
            {
                _options = settingsProtectedFlag.SetBoolean(_options, value);
                //if (value == true)
                //{
                //    field_1_protect = 1;
                //}
                //else
                //{
                //    field_1_protect = 0;
                //}
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[WINDOWPROTECT]\n");
            buffer.Append("    .protect         = ").Append(Protect)
                .Append("\n");
            buffer.Append("[/WINDOWPROTECT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(_options);
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