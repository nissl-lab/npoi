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

namespace NPOI.SS.Formula.PTG
{

    using System;
    using System.Text;
    using NPOI.Util;
    


    /**
     *
     * @author  andy
     * @author Jason Height (jheight at chariot dot net dot au)
     * @author dmui (save existing implementation)
     */
    public class ExpPtg : ControlPtg
    {
        private const int SIZE = 5;
        public const byte sid = 0x1;
        private short field_1_first_row;
        private short field_2_first_col;

        public ExpPtg(ILittleEndianInput in1)
        {
            field_1_first_row = in1.ReadShort();
            field_2_first_col = in1.ReadShort();
        }

        public ExpPtg(int firstRow, int firstCol)
        {
            this.field_1_first_row = (short)firstRow;
            this.field_2_first_col = (short)firstCol;
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(field_1_first_row);
            out1.WriteShort(field_2_first_col);
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public short Row
        {
            get { return field_1_first_row; }
        }

        public short Column
        {
            get { return field_2_first_col; }
        }

        public override String ToFormulaString()
        {
            throw new RecordFormatException("Coding Error: Expected ExpPtg to be Converted from Shared to Non-Shared Formula by ValueRecordsAggregate, but it wasn't");
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder("[Array Formula or Shared Formula]\n");
            buffer.Append("row = ").Append(Row).Append("\n");
            buffer.Append("col = ").Append(Column).Append("\n");
            return buffer.ToString();
        }
    }
}