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

    /// <summary>
    /// <para>
    /// This class is a Container for POI usermodel row=0 column=0 cell references.
    /// It is barely a Container for these two coordinates. The implementation
    /// of the Comparable interface sorts by "natural" order top left to bottom right.
    /// </para>
    /// <para>
    /// Use <tt>CellAddress</tt> when you want to refer to the location of a cell in a sheet
    /// when the concept of relative/absolute does not apply (such as the anchor location
    /// of a cell comment). Use <see cref="CellReference"/> when the concept of
    /// relative/absolute does apply (such as a cell reference in a formula).
    /// <tt>CellAddress</tt>es do not have a concept of "sheet", while <tt>CellReference</tt>s do.
    /// </para>
    /// </summary>
    public class CellAddress : IComparable<CellAddress>
    {
        /// <summary>
        /// A constant for references to the first cell in a sheet. */
        /// </summary>
        public static CellAddress A1 = new CellAddress(0, 0);

        private int _row;
        private int _col;

        /// <summary>
        /// Create a new CellAddress object.
        /// </summary>
        /// <param name="row">Row index (first row is 0)</param>
        /// <param name="column">Column index (first column is 0)</param>
        public CellAddress(int row, int column)
                : base()
        {

            this._row = row;
            this._col = column;
        }

        /// <summary>
        /// Create a new CellAddress object.
        /// </summary>
        /// <param name="Address">a cell Address in A1 format. Address may not contain sheet name or dollar signs.
        /// (that is, Address is not a cell reference. Use <see cref="CellAddress(CellReference)" /> instead if
        /// starting with a cell reference.)
        /// </param>
        public CellAddress(String address)
        {
            int length = address.Length;

            int loc = 0;
            // step over column name chars until first digit for row number.
            for(; loc < length; loc++)
            {
                char ch = address[loc];
                if(Char.IsDigit(ch))
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

        /// <summary>
        /// Create a new CellAddress object.
        /// </summary>
        /// <param name="reference">a reference to a cell</param>
        public CellAddress(CellReference reference)
                : this(reference.Row, reference.Col)
        {

        }

        /// <summary>
        /// Create a new CellAddress object
        /// </summary>
        /// <param name="address">a CellAddress</param>
        public CellAddress(CellAddress address)
            : this(address.Row, address.Column)
        {
        }

        /// <summary>
        /// Create a new CellAddress object.
        /// </summary>
        /// <param name="cell">the Cell to Get the location of</param>
        public CellAddress(ICell cell)
                : this(cell.RowIndex, cell.ColumnIndex)
        {

        }

        /// <summary>
        /// Get the cell Address row
        /// </summary>
        /// <return>row</return>
        public int Row
        {
            get
            {
                return _row;
            }
        }

        /// <summary>
        /// Get the cell Address column
        /// </summary>
        /// <return>column</return>
        public int Column
        {
            get
            {
                return _col;
            }
        }

        /// <summary>
        /// Compare this CellAddress using the "natural" row-major, column-minor ordering.
        /// That is, top-left to bottom-right ordering.
        /// </summary>
        /// <param name="other">other</param>
        /// <return><list type="bullet">
        /// <item><description>-1 if this CellAddress is before (above/left) of other</li></description></item>
        /// <item><description>0 if Addresses are the same</li></description></item>
        /// <item><description>1 if this CellAddress is After (below/right) of other</li></description></item>
        /// </list>
        /// </return>

        public int CompareTo(CellAddress other)
        {
            int r = this._row - other._row;
            if(r != 0)
                return r;

            r = this._col - other._col;
            if(r != 0)
                return r;

            return 0;
        }


        public override bool Equals(Object o)
        {
            if(this == o)
            {
                return true;
            }
            if(!(o is CellAddress))
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

        /// <summary>
        /// Same as <see cref="ToString()" />
        /// </summary>
        /// <return>A1-style cell Address string representation</return>
        public String FormatAsString()
        {
            return CellReference.ConvertNumToColString(this._col) + (this._row + 1);
        }
    }
}
