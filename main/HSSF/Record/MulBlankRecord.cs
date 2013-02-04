
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
        

/*
 * MulBlankRecord.java
 *
 * Created on December 10, 2001, 12:49 PM
 */
namespace NPOI.HSSF.Record
{
    using NPOI.Util;
    using System;
    using System.Text;

    /**
     * Title:        Mulitple Blank cell record 
     * Description:  Represents a  Set of columns in a row with no value but with styling.
     *               In this release we have Read-only support for this record type.
     *               The RecordFactory Converts this to a Set of BlankRecord objects.
     * REFERENCE:  PG 329 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Glen Stampoultzis (glens at apache.org)
     * @version 2.0-pre
     * @see org.apache.poi.hssf.record.BlankRecord
     */

    public class MulBlankRecord : StandardRecord
    {
        public const short sid = 0xbe;
        //private short             field_1_row;
        private int _row;
        private int _first_col;
        private short[] _xfs;
        private int _last_col;

        /** Creates new MulBlankRecord */

        public MulBlankRecord()
        {
        }

        public MulBlankRecord(int row, int firstCol, short[] xfs)
        {
            _row = row;
            _first_col = firstCol;
            _xfs = xfs;
            _last_col = firstCol + xfs.Length - 1;
        }

        /**
         * Constructs a MulBlank record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public MulBlankRecord(RecordInputStream in1)
        {
            _row = in1.ReadUShort();
            _first_col = in1.ReadShort();
            _xfs = ParseXFs(in1);
            _last_col = in1.ReadShort();
        }

        /**
         * Get the row number of the cells this represents
         *
         * @return row number
         */

        //public short Row
        public int Row
        {
            get { return _row; }
        }

        /**
         * starting column (first cell this holds in the row)
         * @return first column number
         */

        public int FirstColumn
        {
            get { return _first_col; }
        }

        /**
         * ending column (last cell this holds in the row)
         * @return first column number
         */

        public int LastColumn
        {
            get { return _last_col; }
        }

        /**
         * Get the number of columns this Contains (last-first +1)
         * @return number of columns (last - first +1)
         */

        public int NumColumns
        {
            get { return _last_col - _first_col + 1; }
        }

        /**
         * returns the xf index for column (coffset = column - field_2_first_col)
         * @param coffset  the column (coffset = column - field_2_first_col)
         * @return the XF index for the column
         */

        public short GetXFAt(int coffset)
        {
            return _xfs[coffset];
        }

        private short[] ParseXFs(RecordInputStream in1)
        {
            short[] retval = new short[(in1.Remaining - 2) / 2];

            for (int idx = 0; idx < retval.Length; idx++)
            {
                retval[idx] = in1.ReadShort();
            }
            return retval;
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[MULBLANK]\n");
            buffer.Append("row  = ")
                .Append(StringUtil.ToHexString(Row)).Append("\n");
            buffer.Append("firstcol  = ")
                .Append(StringUtil.ToHexString(FirstColumn)).Append("\n");
            buffer.Append(" lastcol  = ")
                .Append(StringUtil.ToHexString(LastColumn)).Append("\n");
            for (int k = 0; k < NumColumns; k++)
            {
                buffer.Append("xf").Append(k).Append("        = ")
                    .Append(StringUtil.ToHexString(GetXFAt(k))).Append("\n");
            }
            buffer.Append("[/MULBLANK]\n");
            return buffer.ToString();
        }

        public override short Sid
        {
            get { return sid; }
        }
        protected override int DataSize
        {
            get { return 6 + _xfs.Length * 2; }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(_row);
            out1.WriteShort(_first_col);
            int nItems = _xfs.Length;
            for (int i = 0; i < nItems; i++)
            {
                out1.WriteShort(_xfs[i]);
            }
            out1.WriteShort(_last_col);
        }
        //poi bug 46776
        public override Object Clone()
        {
            // immutable - so OK to return this
            return this;
        }
    }
}