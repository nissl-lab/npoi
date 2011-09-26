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
    using System;


    /// <summary>
    /// A client anchor Is attached to an excel worksheet.  It anchors against a
    /// top-left and buttom-right cell.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class HSSFClientAnchor : HSSFAnchor, NPOI.SS.UserModel.IClientAnchor
    {
        int col1;
        int row1;
        int col2;
        int row2;
        int anchorType;

        /// <summary>
        /// Creates a new client anchor and defaults all the anchor positions to 0.
        /// </summary>
        public HSSFClientAnchor()
        {
        }

        /// <summary>
        /// Creates a new client anchor and Sets the top-left and bottom-right
        /// coordinates of the anchor.
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

            this.col1 = col1;
            this.row1 = row1;
            this.col2 = col2;
            this.row2 = row2;
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
            get { return col1; }
            set
            {
                CheckRange(value, 0, 255, "col1");
                this.col1 = value;
            }
        }

        /// <summary>
        /// Gets or sets the col2.
        /// </summary>
        /// <value>The col2.</value>
        public int Col2
        {
            get { return col2; }
            set
            {
                CheckRange(value, 0, 255, "col2");
                this.col2 = value;
            }
        }

        /// <summary>
        /// Gets or sets the row1.
        /// </summary>
        /// <value>The row1.</value>
        public int Row1
        {
            get { return row1; }
            set
            {
                CheckRange(value, 0, 256 * 256, "row1");
                this.row1 = value;
            }
        }

        /// <summary>
        /// Gets or sets the row2.
        /// </summary>
        /// <value>The row2.</value>
        public int Row2
        {
            get { return row2; }
            set
            {
                CheckRange(value, 0, 256 * 256, "row2");
                this.row2 = value;
            }
        }

        /// <summary>
        /// Sets the top-left and bottom-right
        /// coordinates of the anchor
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

            this.col1 = col1;
            this.row1 = row1;
            this.Dx1 = x1;
            this.Dy1 = y1;
            this.col2 = col2;
            this.row2 = row2;
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
                if (col1 == col2)
                    return Dx1 > Dx2;
                else
                    return col1 > col2;
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
                if (row1 == row2)
                    return Dy1 > Dy2;
                else
                    return row1 > row2;
            }
        }

        /// <summary>
        /// Gets the anchor type
        /// 0 = Move and size with Cells, 2 = Move but don't size with cells, 3 = Don't move or size with cells.
        /// </summary>
        /// <value>The type of the anchor.</value>
        public int AnchorType
        {
            get{return anchorType;}
            set { this.anchorType = value; }
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
                throw new ArgumentException(varName + " must be between " + minRange + " and " + maxRange);
        }
    }
}