
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
     * Title: Protect Record
     * Description:  defines whether a sheet or workbook is protected (HSSF DOES NOT SUPPORT ENCRYPTION)
     * (kindly ask the US government to stop having arcane stupid encryption laws and we'll support it) 
     * (after all terrorists will all use US-legal encrypton right??)
     * HSSF now supports the simple "protected" sheets (where they are not encrypted and open office et al
     * ignore the password record entirely).
     * REFERENCE:  PG 373 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     */

    public class ProtectRecord
       : StandardRecord
    {
        public const short sid = 0x12;
        private static BitField protectFlag = BitFieldFactory.GetInstance(0x0001);
        private short _options;
        public ProtectRecord(short options)
        {
            _options = options;
        }

        /**
         * Constructs a Protect record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public ProtectRecord(RecordInputStream in1)
            : this(in1.ReadShort())
        {

        }
        public ProtectRecord(bool isProtected)
            : this(0)
        {
            this.Protect = (isProtected);
        }
        /**
         * Get whether the sheet is protected or not
         * @return whether to protect the sheet or not
         */

        public bool Protect
        {
            get { return protectFlag.IsSet(_options); }
            set
            {
                _options = (short)protectFlag.SetBoolean(_options, value);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PROTECT]\n");
            buffer.Append("    .options = ").Append(HexDump.ShortToHex(_options))
                .Append("\n");
            buffer.Append("[/PROTECT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(_options);
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            return new ProtectRecord(_options);
        }
    }
}