
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
     * Title:        Protection Revision 4 Record
     * Description:  describes whether this is a protected shared/tracked workbook
     *  ( HSSF does not support encryption because we don't feel like going to jail ) 
     * REFERENCE:  PG 373 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class ProtectionRev4Record
       : StandardRecord
    {
        public const short sid = 0x1af;

        private static BitField protectedFlag = BitFieldFactory.GetInstance(0x0001);
        private short _options;

        public ProtectionRev4Record(short options)
        {
            _options = options;
        }
        public ProtectionRev4Record(bool protect):this(0)
        {
            Protect = protect;
        }
       
        /**
         * Constructs a ProtectionRev4 record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public ProtectionRev4Record(RecordInputStream in1):
            this(in1.ReadShort())
        {
            
        }

        /**
         * Get whether the this is protected shared/tracked workbook or not
         * @return whether to protect the workbook or not
         */

        public bool Protect
        {
            get { return protectedFlag.IsSet(_options); }
            set
            {
                _options=(short)protectedFlag.SetBoolean(_options, value);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PROT4REV]\n");
            buffer.Append("    .protect         = ").Append(Protect)
                .Append("\n");
            buffer.Append("[/PROT4REV]\n");
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
    }

}