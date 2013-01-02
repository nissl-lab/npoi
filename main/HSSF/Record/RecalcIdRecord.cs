
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
     * Title: Recalc Id Record
     * Description:  This record Contains an ID that marks when a worksheet was last
     *               recalculated. It's an optimization Excel uses to determine if it
     *               needs to  recalculate the spReadsheet when it's opened. So far, only
     *               the two values <c>0xC1 0x01 0x00 0x00 0x80 0x38 0x01 0x00</c>
     *               (do not recalculate) and <c>0xC1 0x01 0x00 0x00 0x60 0x69 0x01
     *               0x00</c> have been seen. If the field <c>isNeeded</c> Is
     *               Set to false (default), then this record Is swallowed during the
     *               serialization Process
     * REFERENCE:  http://chicago.sourceforge.net/devel/docs/excel/biff8.html
     * @author Luc Girardin (luc dot girardin at macrofocus dot com)
     * @version 2.0-pre
     * @see org.apache.poi.hssf.model.Workbook
     */

    public class RecalcIdRecord
       : StandardRecord
    {
        public const short sid = 0x1c1;
        //public short[] field_1_recalcids;

        //private bool isNeeded = true;
        private int _reserved0;
        //private bool isNeeded = true;
        /**
     * An unsigned integer that specifies the recalculation engine identifier
     * of the recalculation engine that performed the last recalculation.
     * If the value is less than the recalculation engine identifier associated with the application,
     * the application will recalculate the results of all formulas on
     * this workbook immediately after loading the file
     */
        private int _engineId;
        public RecalcIdRecord()
        {
            _reserved0 = 0;
            _engineId = 0;
        }

        /**
         * Constructs a RECALCID record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public RecalcIdRecord(RecordInputStream in1)
        {
            in1.ReadUShort(); // field 'rt' should have value 0x01C1, but Excel doesn't care during reading
    	_reserved0 = in1.ReadUShort();
    	_engineId = in1.ReadInt();
        }

        // /**
        // * Set the recalc array.
        // * @param array of recalc id's
        // */

        //public void SetRecalcIdArray(short[] array)
        //{
        //    field_1_recalcids = array;
        //}

        // /**
        // * Get the recalc array.
        // * @return array of recalc id's
        // */

        //public short[] GetRecalcIdArray()
        //{
        //    return field_1_recalcids;
        //}

        public bool IsNeeded
        {
            get { return  true; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();
            buffer.Append("[RECALCID]\n");
            buffer.Append("    .reserved = ").Append(HexDump.ShortToHex(_reserved0)).Append("\n");
            buffer.Append("    .engineId = ").Append(HexDump.IntToHex(_engineId)).Append("\n");
            buffer.Append("[/RECALCID]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid); // always write 'rt' field as 0x01C1
            out1.WriteShort(_reserved0);
            out1.WriteInt(_engineId);
        }
        public int EngineId
        {
            set
            {
                _engineId = value;
            }
            get
            {
                return _engineId;
            }
        }
        protected override int DataSize
        {
            get
            {
                return 8;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }
    }
}
