/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.SS.Util
{
    using System;



    using NPOI.SS.UserModel;

    /**
     * <p>This class is a Container for POI usermodel row=0 column=0 cell references.
     * It is barely a Container for these two coordinates. The implementation
     * of the Comparable interface sorts by "natural" order top left to bottom right.</p>
     * 
     * <p>Use <tt>CellAddress</tt> when you want to refer to the location of a cell in a sheet
     * when the concept of relative/absolute does not apply (such as the anchor location 
     * of a cell comment). Use {@link CellReference} when the concept of
     * relative/absolute does apply (such as a cell reference in a formula).
     * <tt>CellAddress</tt>es do not have a concept of "sheet", while <tt>CellReference</tt>s do.</p>
     */
    public class CellAddress : IComparable<CellAddress>
    {
        /** A constant for references to the first cell in a sheet. */
        public static CellAddress A1 = new CellAddress(0, 0);

        private int _row;
        private int _col;

        /**
         * Create a new CellAddress object.
         *
         * @param row Row index (first row is 0)
         * @param column Column index (first column is 0)
         */
        public CellAddress(int row, int column)
                : base()
        {

            this._row = row;
            this._col = column;
        }

        /**
         * Create a new CellAddress object.
         *
         * @param Address a cell Address in A1 format. Address may not contain sheet name or dollar signs.
         * (that is, Address is not a cell reference. Use {@link #CellAddress(CellReference)} instead if
         * starting with a cell reference.)
         */
        public CellAddress(String address)
        {
            int length = address.Length;

            int loc = 0;
            // step over column name chars until first digit for row number.
            for (; loc < length; loc++)
            {
                char ch = address[loc];
                if (Char.IsDigit(ch))
                {
                    break;
                }
            }

            ReadOnlySpan<char> sCol = address.AsSpan(0, loc);
            ReadOnlySpan<char> sRow = address.AsSpan(loc);

            // FIXME: breaks if Address Contains a sheet name or dollar signs from an absolute CellReference
            CellReferenceParser.TryParsePositiveInt32Fast(sRow, out var rowNumber);
            this._row = rowNumber - 1;
            this._col = CellReference.ConvertColStringToIndex(sCol);
        }

        /**
         * Create a new CellAddress object.
         *
         * @param reference a reference to a cell
         */
        public CellAddress(CellReference reference)
                : this(reference.Row, reference.Col)
        {

        }

        /**
         * Create a new CellAddress object.
         *
         * @param cell the Cell to Get the location of
         */
        public CellAddress(ICell cell)
                : this(cell.RowIndex, cell.ColumnIndex)
        {

        }

        /**
         * Get the cell Address row
         *
         * @return row
         */
        public int Row
        {
            get
            {
                return _row;
            }
        }

        /**
         * Get the cell Address column
         *
         * @return column
         */
        public int Column
        {
            get
            {
                return _col;
            }
        }

        /**
         * Compare this CellAddress using the "natural" row-major, column-minor ordering.
         * That is, top-left to bottom-right ordering.
         * 
         * @param other
         * @return <ul>
         * <li>-1 if this CellAddress is before (above/left) of other</li>
         * <li>0 if Addresses are the same</li>
         * <li>1 if this CellAddress is After (below/right) of other</li>
         * </ul>
         */

        public int CompareTo(CellAddress other)
        {
            int r = this._row - other._row;
            if (r != 0) return r;

            r = this._col - other._col;
            if (r != 0) return r;

            return 0;
        }


        public override bool Equals(Object o)
        {
            if (this == o)
            {
                return true;
            }
            if (!(o is CellAddress))
            {
                return false;
            }

            CellAddress other = (CellAddress)o;
            return _row == other._row
                    && _col == other._col
            ;
        }


        public override int GetHashCode()
        {
            return this._row + this._col << 16;
        }


        public override String ToString()
        {
            return FormatAsString();
        }

        /**
         * Same as {@link #ToString()}
         * @return A1-style cell Address string representation
         */
        public String FormatAsString()
        {
            return CellReference.ConvertNumToColString(this._col) + (this._row + 1);
        }
    }
}
