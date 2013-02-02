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
    using NPOI.Util;
    
    using NPOI.SS.Util;


    /**
     * ReferencePtgBase - handles references (such as A1, A2, IA4)
     * @author  Andrew C. Oliver (acoliver@apache.org)
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    [Serializable]
    public abstract class RefPtgBase : OperandPtg
    {
        /** The row index - zero based Unsigned 16 bit value */
        private int field_1_row;
        /** Field 2
         * - lower 8 bits is the zero based Unsigned byte column index
         * - bit 16 - IsRowRelative
         * - bit 15 - IsColumnRelative
         */
        private int field_2_col;
        private static BitField rowRelative = BitFieldFactory.GetInstance(0x8000);
        private static BitField colRelative = BitFieldFactory.GetInstance(0x4000);
        private static BitField column = BitFieldFactory.GetInstance(0x3FFF);

        protected RefPtgBase()
        {
            //Required for Clone methods
        }

        /**
         * Takes in a String representation of a cell reference and Fills out the
         * numeric fields.
         */
        protected RefPtgBase(String cellref)
        {
            CellReference c = new CellReference(cellref);
            Row = c.Row;
            Column = c.Col;
            IsColRelative = !c.IsColAbsolute;
            IsRowRelative = !c.IsRowAbsolute;
        }
        protected RefPtgBase(CellReference c)
        {
            Row = (c.Row);
            Column = (c.Col);
            IsColRelative = (!c.IsColAbsolute);
            IsRowRelative = (!c.IsRowAbsolute);
        }
        protected RefPtgBase(int row, int column, bool isRowRelative, bool isColumnRelative)
        {
            this.Row = row;
            this.Column = column;
            this.IsRowRelative = isRowRelative;
            this.IsColRelative = isColumnRelative;
        }

        protected RefPtgBase(ILittleEndianInput in1)
        {
            field_1_row = in1.ReadUShort();
            field_2_col = in1.ReadUShort();
        }
        protected void ReadCoordinates(ILittleEndianInput in1)
        {
            field_1_row = in1.ReadUShort();
            field_2_col = in1.ReadUShort();
        }
        protected void WriteCoordinates(ILittleEndianOutput out1)
        {
            out1.WriteShort(field_1_row);
            out1.WriteShort(field_2_col);
        }

        protected void WriteCoordinates(byte[] array, int offset)
        {
            LittleEndian.PutUShort(array, offset + 0, field_1_row);
            LittleEndian.PutUShort(array, offset + 2, field_2_col);
        }
        /**
         * Returns the row number as a short, which will be
         *  wrapped (negative) for values between 32769 and 65535
         */
        public int Row
        {
            get { return field_1_row; }
            set
            {
                field_1_row = value;
            }
        }
        /**
         * Returns the row number as an int, between 0 and 65535
         */
        public int RowAsInt
        {
            get { return field_1_row; }
        }

        public bool IsRowRelative
        {
            get { return rowRelative.IsSet(field_2_col); }
            set
            {
                field_2_col = rowRelative.SetBoolean(field_2_col, value);
            }
        }

        public bool IsColRelative
        {
            get { return colRelative.IsSet(field_2_col); }
            set
            {
                field_2_col = colRelative.SetBoolean(field_2_col, value);
            }
        }

        public int ColumnRawX
        {
            get { return field_2_col; }
            set { field_2_col = value; }
        }

        public int Column
        {
            get { return column.GetValue(field_2_col); }
            set
            {
                /*
                 if (value < 0 || value >= 0x100)
                {
                    throw new ArgumentException("Specified colIx (" + value + ") is out of range");
                }*/
                field_2_col = column.SetValue(field_2_col, value);
            }
        }
        public String FormatReferenceAsString()
        {
            // Only make cell references as needed. Memory is an issue
            CellReference cr = new CellReference(Row, Column, !IsRowRelative, !IsColRelative);
            return cr.FormatAsString();
        }


        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_REF; }
        }
    }
}