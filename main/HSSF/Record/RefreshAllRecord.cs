
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
       : StandardRecord
    {
        public const short sid = 0x1B7;
        //private short field_1_refreshall;
        private static BitField refreshFlag = BitFieldFactory.GetInstance(0x0001);

    private int _options;
    public RefreshAllRecord(int options)
    {
        _options = options;
        }

        /**
         * Constructs a RefreshAll record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public RefreshAllRecord(RecordInputStream in1)
            : this(in1.ReadUShort())
        {
        }
        public RefreshAllRecord(bool refreshAll)
            : this(0)
        {
            RefreshAll = (refreshAll);
        }
        /**
         * Get whether to refresh all external data when loading a sheet
         * @return refreshall or not
         */

        public bool RefreshAll
        {
            get { return refreshFlag.IsSet(_options); }
            set
            {
                _options = refreshFlag.SetBoolean(_options, value);
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
        public override Object Clone()
        {
            return new RefreshAllRecord(_options);
        }
    }
}