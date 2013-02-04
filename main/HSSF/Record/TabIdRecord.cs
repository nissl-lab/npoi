
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
     * Title: Sheet Tab Index Array Record
     * Description:  Contains an array of sheet id's.  Sheets always keep their ID
     *               regardless of what their name Is.
     * REFERENCE:  PG 412 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @version 2.0-pre
     */

    public class TabIdRecord
       : StandardRecord
    {
        public const short sid = 0x13d;
        private static short[] EMPTY_SHORT_ARRAY = { };
        public short[] _tabids;

        public TabIdRecord()
        {
            _tabids = EMPTY_SHORT_ARRAY;
        }

        /**
         * Constructs a TabID record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public TabIdRecord(RecordInputStream in1)
        {
            _tabids = new short[in1.Remaining / 2];
            for (int k = 0; k < _tabids.Length; k++)
            {
                _tabids[k] = in1.ReadShort();
            }
        }

        /**
         * Set the tab array.  (0,1,2).
         * @param array of tab id's {0,1,2}
         */

        public void SetTabIdArray(short[] array)
        {
            _tabids = array;
        }

        /**
         * Get the tab array.  (0,1,2).
         * @return array of tab id's {0,1,2}
         */

        //public short[] GetTabIdArray()
        //{
        //    return field_1_tabids;
        //}

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[TABID]\n");
            buffer.Append("    .elements        = ").Append(_tabids.Length)
                .Append("\n");
            for (int k = 0; k < _tabids.Length; k++)
            {
                buffer.Append("    .element_" + k + "       = ")
                    .Append(_tabids[k]).Append("\n");
            }
            buffer.Append("[/TABID]\n");
            return buffer.ToString();
        }

        //public override int Serialize(int offset, byte [] data)
        //{
        //    short[] tabids = GetTabIdArray();
        //    short Length = (short)(tabids.Length * 2);
        //    int byteoffset = 4;

        //    LittleEndian.PutShort(data, 0 + offset, sid);
        //    LittleEndian.PutShort(data, 2 + offset,
        //                          ((short)Length));   // nubmer tabids *

        //    // 2 (num bytes in a short)
        //    for (int k = 0; k < (Length / 2); k++)
        //    {
        //        LittleEndian.PutShort(data, byteoffset + offset, tabids[k]);
        //        byteoffset += 2;
        //    }
        //    return RecordSize;
        //}

        //public override int RecordSize
        //{
        //    get { return 4 + (GetTabIdArray().Length * 2); }
        //}
        public override void Serialize(ILittleEndianOutput out1)
        {
            short[] tabids = _tabids;

            for (int i = 0; i < tabids.Length; i++)
            {
                out1.WriteShort(tabids[i]);
            }
        }

        protected override int DataSize
        {
            get
            {
                return _tabids.Length * 2;
            }
        }
        public override short Sid
        {
            get { return sid; }
        }
    }
}
