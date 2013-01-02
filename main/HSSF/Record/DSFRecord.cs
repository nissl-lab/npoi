
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

    public class DSFRecord: StandardRecord
    {
        public const short sid = 0x161;
        private int _options;
        private static BitField biff5BookStreamFlag = BitFieldFactory.GetInstance(0x0001);

        private DSFRecord(int options)
        {
            _options = options;
        }
        public DSFRecord(bool isBiff5BookStreamPresent):this(0)
        {
            
            _options = biff5BookStreamFlag.SetBoolean(0, isBiff5BookStreamPresent);
        }
        /**
         * Constructs a DBCellRecord and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public DSFRecord(RecordInputStream in1):this(in1.ReadShort())
        {
            
        }

        public bool IsBiff5BookStreamPresent
        {
          get
            {
            return biff5BookStreamFlag.IsSet(_options);
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[DSF]\n");
            buffer.Append("    .IsDSF           = ")
                .Append(StringUtil.ToHexString(_options)).Append("\n");
            buffer.Append("[/DSF]\n");
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
