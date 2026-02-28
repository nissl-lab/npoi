/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula.PTG
{
    using System;
    using System.Text;
    using NPOI.Util;


    /**
     * This ptg indicates a data table.
     * It only occurs in a FORMULA record, never in an
     *  ARRAY or NAME record.  When ptgTbl occurs in a
     *  formula, it is the only token in the formula.
     *
     * This indicates that the cell containing the
     *  formula is an interior cell in a data table;
     *  the table description is found in a TABLE
     *  record. Rows and columns which contain input
     *  values to be substituted in the table do
     *  not contain ptgTbl.
     * See page 811 of the june 08 binary docs.
     */
    public class TblPtg : ControlPtg
    {
        private const int SIZE = 5;
        public const byte sid = 0x02;
        /** The row number of the upper left corner */
        private int field_1_first_row;
        /** The column number of the upper left corner */
        private int field_2_first_col;

        public TblPtg(ILittleEndianInput in1)
        {
            field_1_first_row = in1.ReadUShort();
            field_2_first_col = in1.ReadUShort();
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteShort(field_1_first_row);
            out1.WriteShort(field_2_first_col);
        }


        public override int Size
        {
            get
            {
                return SIZE;
            }
        }

        public int Row
        {
            get
            {
                return field_1_first_row;
            }
        }

        public int Column
        {
            get
            {
                return field_2_first_col;
            }
        }

        public override String ToFormulaString()
        {
            // table(....)[][]
            throw new RecordFormatException("Table and Arrays are not yet supported");
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder("[Data Table - Parent cell is an interior cell in a data table]\n");
            buffer.Append("top left row = ").Append(Row).Append("\n");
            buffer.Append("top left col = ").Append(Column).Append("\n");
            return buffer.ToString();
        }
    }
}