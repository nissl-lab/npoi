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


namespace NPOI.HSSF.UserModel
{
    using NPOI.DDF;
    using NPOI.SS.UserModel;
    using System;


    /// <summary>
    /// A client anchor Is attached to an excel worksheet.  It anchors against a
    /// top-left and buttom-right cell.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class HSSFClientAnchor : HSSFAnchor, IClientAnchor
    {
        private EscherClientAnchorRecord _escherClientAnchor;
        public HSSFClientAnchor(EscherClientAnchorRecord escherClientAnchorRecord)
        {
            this._escherClientAnchor = escherClientAnchorRecord;
        }
        /// <summary>
        /// Creates a new client anchor and defaults all the anchor positions to 0.
        /// </summary>
        public HSSFClientAnchor()
        {
            //Is this necessary?
            this._escherClientAnchor = new EscherClientAnchorRecord();
        }

        /// <summary>
        /// Creates a new client anchor and Sets the top-left and bottom-right
        /// coordinates of the anchor.
        /// 
        /// Note: Microsoft Excel seems to sometimes disallow 
        /// higher y1 than y2 or higher x1 than x2 in the anchor, you might need to 
        /// reverse them and draw shapes vertically or horizontally flipped! 
        /// </summary>
        /// <param name="dx1">the x coordinate within the first cell.</param>
        /// <param name="dy1">the y coordinate within the first cell.</param>
        /// <param name="dx2">the x coordinate within the second cell.</param>
        /// <param name="dy2">the y coordinate within the second cell.</param>
        /// <param name="col1">the column (0 based) of the first cell.</param>
        /// <param name="row1">the row (0 based) of the first cell.</param>
        /// <param name="col2">the column (0 based) of the second cell.</param>
        /// <param name="row2">the row (0 based) of the second cell.</param>
        public HSSFClientAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
            : base(dx1, dy1, dx2, dy2)
        {


            CheckRange(dx1, 0, 1023, "dx1");
            CheckRange(dx2, 0, 1023, "dx2");
            CheckRange(dy1, 0, 255, "dy1");
            CheckRange(dy2, 0, 255, "dy2");
            CheckRange(col1, 0, 255, "col1");
            CheckRange(col2, 0, 255, "col2");
            CheckRange(row1, 0, 255 * 256, "row1");
            CheckRange(row2, 0, 255 * 256, "row2");

            Col1=((short)Math.Min(col1, col2));
            Col2=((short)Math.Max(col1, col2));
            Row1=((short)Math.Min(row1, row2));
            Row2=((short)Math.Max(row1, row2));

            if (col1 > col2)
            {
                _isHorizontallyFlipped = true;
            }
            if (row1 > row2)
            {
                _isVerticallyFlipped = true;
            }
        }

        /// <summary>
        /// Calculates the height of a client anchor in points.
        /// </summary>
        /// <param name="sheet">the sheet the anchor will be attached to</param>
        /// <returns>the shape height.</returns>     
        public float GetAnchorHeightInPoints(NPOI.SS.UserModel.ISheet sheet)
        {
            int y1 = Dy1;
            int y2 = Dy2;
            int row1 = Math.Min(Row1, Row2);
            int row2 = Math.Max(Row1, Row2);

            float points = 0;
            if (row1 == row2)
            {
                points = ((y2 - y1) / 256.0f) * GetRowHeightInPoints(sheet, row2);
            }
            else
            {
                points += ((256.0f - y1) / 256.0f) * GetRowHeightInPoints(sheet, row1);
                for (int i = row1 + 1; i < row2; i++)
                {
                    points += GetRowHeightInPoints(sheet, i);
                }
                points += (y2 / 256.0f) * GetRowHeightInPoints(sheet, row2);
            }

            return points;
        }

        /// <summary>
        /// Gets the row height in points.
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="rowNum">The row num.</param>
        /// <returns></returns>
        private float GetRowHeightInPoints(NPOI.SS.UserModel.ISheet sheet, int rowNum)
        {
            NPOI.SS.UserModel.IRow row = sheet.GetRow(rowNum);
            if (row == null)
                return sheet.DefaultRowHeightInPoints;
            else
                return row.HeightInPoints;
        }

        /// <summary>
        /// Gets or sets the col1.
        /// </summary>
        /// <value>The col1.</value>
        public int Col1
        {
            get { return _escherClientAnchor.Col1; }
            set
            {
                CheckRange(value, 0, 255, "col1");
                _escherClientAnchor.Col1 = (short)value;
            }
        }

        /// <summary>
        /// Gets or sets the col2.
        /// </summary>
        /// <value>The col2.</value>
        public int Col2
        {
            get { return _escherClientAnchor.Col2; }
            set
            {
                CheckRange(value, 0, 255, "col2");
                _escherClientAnchor.Col2 = (short)value;
            }
        }

        /// <summary>
        /// Gets or sets the row1.
        /// </summary>
        /// <value>The row1.</value>
        public int Row1
        {
            get { return _escherClientAnchor.Row1; }
            set
            {
                CheckRange(value, 0, 256 * 256, "row1");
                _escherClientAnchor.Row1 = (short)value; 
            }
        }

        /// <summary>
        /// Gets or sets the row2.
        /// </summary>
        /// <value>The row2.</value>
        public int Row2
        {
            get { return _escherClientAnchor.Row2; }
            set
            {
                CheckRange(value, 0, 256 * 256, "row2");
                _escherClientAnchor.Row2 = (short)value;
            }
        }

        /// <summary>
        /// Sets the top-left and bottom-right
        /// coordinates of the anchor
        /// 
        /// Note: Microsoft Excel seems to sometimes disallow 
        /// higher y1 than y2 or higher x1 than x2 in the anchor, you might need to 
        /// reverse them and draw shapes vertically or horizontally flipped! 
        /// </summary>
        /// <param name="col1">the column (0 based) of the first cell.</param>
        /// <param name="row1"> the row (0 based) of the first cell.</param>
        /// <param name="x1">the x coordinate within the first cell.</param>
        /// <param name="y1">the y coordinate within the first cell.</param>
        /// <param name="col2">the column (0 based) of the second cell.</param>
        /// <param name="row2">the row (0 based) of the second cell.</param>
        /// <param name="x2">the x coordinate within the second cell.</param>
        /// <param name="y2">the y coordinate within the second cell.</param>
        public void SetAnchor(short col1, int row1, int x1, int y1, short col2, int row2, int x2, int y2)
        {
            CheckRange(x1, 0, 1023, "dx1");
            CheckRange(x2, 0, 1023, "dx2");
            CheckRange(y1, 0, 255, "dy1");
            CheckRange(y2, 0, 255, "dy2");
            CheckRange(col1, 0, 255, "col1");
            CheckRange(col2, 0, 255, "col2");
            CheckRange(row1, 0, 255 * 256, "row1");
            CheckRange(row2, 0, 255 * 256, "row2");

            this.Col1 = col1;
            this.Row1 = row1;
            this.Dx1 = x1;
            this.Dy1 = y1;
            this.Col2 = col2;
            this.Row2 = row2;
            this.Dx2 = x2;
            this.Dy2 = y2;
        }

        /// <summary>
        /// Gets a value indicating whether this instance is horizontally flipped.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the anchor goes from right to left; otherwise, <c>false</c>.
        /// </value>
        public override bool IsHorizontallyFlipped
        {
            get
            {
                return _isHorizontallyFlipped;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is vertically flipped.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the anchor goes from bottom to top.; otherwise, <c>false</c>.
        /// </value>
        public override bool IsVerticallyFlipped
        {
            get
            {
                return _isVerticallyFlipped;
            }
        }

        /// <summary>
        /// Gets the anchor type
        /// 0 = Move and size with Cells, 2 = Move but don't size with cells, 3 = Don't move or size with cells.
        /// </summary>
        /// <value>The type of the anchor.</value>
        public AnchorType AnchorType
        {
            get { return (AnchorType)_escherClientAnchor.Flag; }
            set { this._escherClientAnchor.Flag = (short)value; }
        }


        /// <summary>
        /// Checks the range.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="minRange">The min range.</param>
        /// <param name="maxRange">The max range.</param>
        /// <param name="varName">Name of the variable.</param>
        private void CheckRange(int value, int minRange, int maxRange, String varName)
        {
            if (value < minRange || value > maxRange)
                throw new ArgumentOutOfRangeException(varName + " must be between " + minRange + " and " + maxRange + ", but was: " + value);
        }
        internal override EscherRecord GetEscherAnchor()
        {
            return _escherClientAnchor;
        }

        protected override void CreateEscherAnchor()
        {
            _escherClientAnchor = new EscherClientAnchorRecord();
        }

        public override bool Equals(Object obj)
        {
            if (obj == null)
                return false;
            if (obj == this)
                return true;
            if (obj.GetType() != GetType())
                return false;
            HSSFClientAnchor anchor = (HSSFClientAnchor)obj;

            return anchor.Col1 == Col1 && anchor.Col2 == Col2 && anchor.Dx1 == Dx1
                    && anchor.Dx2 == Dx2 && anchor.Dy1 == Dy1 && anchor.Dy2 == Dy2
                    && anchor.Row1 == Row1 && anchor.Row2 == Row2 && anchor.AnchorType == AnchorType;
        }
        public override int GetHashCode()
        {
            return Col1.GetHashCode() ^ Col2.GetHashCode() ^ Dx1.GetHashCode()
                   ^ Dx2.GetHashCode() ^ Dy1.GetHashCode() ^ Dy2.GetHashCode()
                    ^Row1.GetHashCode() ^  Row2.GetHashCode() ^ AnchorType.GetHashCode();
        }
        public override int Dx1
        {
            get
            {
                return _escherClientAnchor.Dx1;
            }
            set
            {
                _escherClientAnchor.Dx1 = (short)value;
            }
        }
        public override int Dx2
        {
            get
            {
                return _escherClientAnchor.Dx2;
            }
            set
            {
                _escherClientAnchor.Dx2 = (short)value;
            }
        }
        public override int Dy1
        {
            get
            {
                return _escherClientAnchor.Dy1;
            }
            set
            {
                _escherClientAnchor.Dy1 = (short)value;
            }
        }
        public override int Dy2
        {
            get
            {
                return _escherClientAnchor.Dy2;
            }
            set
            {
                _escherClientAnchor.Dy2 = (short)value;
            }
        }
    }
}