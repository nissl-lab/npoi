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

using System;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using SixLabors.Fonts;

namespace NPOI.XSSF.UserModel
{
        /// <summary>
    /// <para>
    /// A client anchor is attached to an excel worksheet.  It anchors against:
    /// <list type="number">
    /// <item><description>A fixed position and fixed size</description></item>
    /// <item><description>A position relative to a cell (top-left) and a fixed size</description></item>
    /// <item><description>A position relative to a cell (top-left) and sized relative to another cell (bottom right)</description></item>
    /// </list>
    /// </para>
    /// <para>
    /// which method is used is determined by the <see cref="AnchorType"/>.
    /// </para>
    /// </summary>
    public class XSSFClientAnchor : XSSFAnchor, IClientAnchor
    {
        /// <summary>
        /// placeholder for zeros when needed for dynamic position calculations
        /// </summary>
        private static CT_Marker EMPTY_MARKER = new CT_Marker();

        private AnchorType anchorType;

        /// <summary>
        /// Starting anchor point (top-left cell + relative offset)
        /// if left null recalculate as needed from point
        /// </summary>
        private CT_Marker cell1;

        /// <summary>
        /// Ending anchor point (bottom-right cell + relative offset)
        /// if left null, re-calculate as needed from size and cell1
        /// </summary>
        private CT_Marker cell2;

        /// <summary>
        /// if present, fixed size of the object to use instead of cell2, which is inferred instead
        /// </summary>
        private CT_PositiveSize2D size;

        /// <summary>
        /// if present, fixed top-left position to use instead of cell1, which is inferred instead
        /// </summary>
        private CT_Point2D position;

        /// <summary>
        /// sheet to base dynamic calculations on, if needed.  Required if size and/or position or Set.
        /// Not needed if cell1/2 are Set explicitly (dynamic sizing and position relative to cells).
        /// </summary>
        private XSSFSheet sheet;

        public int left
        {
            get;
        }
        public int top
        {
            get;
        }
        public int width
        {
            get;
        }
        public int height
        {
            get;
        }

        /// <summary>
        /// Creates a new client anchor and defaults all the anchor positions to 0.
        /// Sets the type to <see cref="AnchorType.MoveAndResize" /> relative to cell range A1:A1.
        /// </summary>
        public XSSFClientAnchor() : this(0, 0, 0, 0, 0, 0, 0, 0)
        {
        }

        /// <summary>
        /// Creates a new client anchor and Sets the top-left and bottom-right
        /// coordinates of the anchor by cell references and offsets.
        /// Sets the type to <see cref="AnchorType.MOVE_AND_RESIZE" />.
        /// </summary>
        /// <param name="dx1"> the x coordinate within the first cell.</param>
        /// <param name="dy1"> the y coordinate within the first cell.</param>
        /// <param name="dx2"> the x coordinate within the second cell.</param>
        /// <param name="dy2"> the y coordinate within the second cell.</param>
        /// <param name="col1">the column (0 based) of the first cell.</param>
        /// <param name="row1">the row (0 based) of the first cell.</param>
        /// <param name="col2">the column (0 based) of the second cell.</param>
        /// <param name="row2">the row (0 based) of the second cell.</param>
        public XSSFClientAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)
        {
            anchorType = AnchorType.MoveAndResize;
            cell1 = new CT_Marker();
            cell1.col = (col1);
            cell1.colOff = (dx1);
            cell1.row = (row1);
            cell1.rowOff = (dy1);
            cell2 = new CT_Marker();
            cell2.col = (col2);
            cell2.colOff = (dx2);
            cell2.row = (row2);
            cell2.rowOff = (dy2);
        }

                /// <summary>
        /// Create XSSFClientAnchor from existing xml beans, sized and positioned relative to a pair of cells.
        /// Sets the type to <see cref="AnchorType.MOVE_AND_RESIZE" />.
        /// </summary>
        /// <param name="cell1">starting anchor point</param>
        /// <param name="cell2">ending anchor point</param>
        internal XSSFClientAnchor(CT_Marker cell1, CT_Marker cell2)
        {
            anchorType = AnchorType.MoveAndResize;
            this.cell1 = cell1;
            this.cell2 = cell2;
        }

        /// <summary>
        /// Create XSSFClientAnchor from existing xml beans, sized and positioned relative to a pair of cells.
        /// Sets the type to <see cref="AnchorType.MoveDontResize" />.
        /// </summary>
        /// <param name="sheet">needed to calculate ending point based on column/row sizes</param>
        /// <param name="cell1">starting anchor point</param>
        /// <param name="size">object size, to calculate ending anchor point</param>
        internal XSSFClientAnchor(XSSFSheet sheet, CT_Marker cell1, CT_PositiveSize2D size)
        {
            anchorType = AnchorType.MoveDontResize;
            this.sheet = sheet;
            this.size = size;
            this.cell1 = cell1;
            //        this.cell2 = calcCell(sheet, cell1, size.Cx, size.Cy);
        }

        /// <summary>
        /// Create XSSFClientAnchor from existing xml beans, sized and positioned relative to a pair of cells.
        /// Sets the type to <see cref="AnchorType.DontMoveAndResize" />.
        /// </summary>
        /// <param name="sheet">needed to calculate starting and ending points based on column/row sizes</param>
        /// <param name="position">starting absolute position</param>
        /// <param name="size">object size, to calculate ending position</param>
        internal XSSFClientAnchor(XSSFSheet sheet, CT_Point2D position, CT_PositiveSize2D size)
        {
            anchorType = AnchorType.DontMoveAndResize;
            this.sheet = sheet;
            this.position = position;
            this.size = size;
            // zeros for row/col/offsets
            //        this.cell1 = calcCell(sheet, EMPTY_MARKER, position.Cx, position.Cy);
            //        this.cell2 = calcCell(sheet, cell1, size.Cx, size.Cy);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="dx1"></param>
        /// <param name="dy1"></param>
        /// <param name="dx2"></param>
        /// <param name="dy2"></param>
        public XSSFClientAnchor(
              ISheet Sheet
            , int dx1
            , int dy1
            , int dx2
            , int dy2
        ) : this() {
            IFont   ift = ((XSSFWorkbook)Sheet.Workbook).GetStylesSource().GetFontAt(0);
            Font    ft  = SheetUtil.IFont2Font(ift);
            var     rt  = TextMeasurer.MeasureSize("0", new TextOptions(ft));
            double  MDW = rt.Width + 1;                                                     //MaximumDigitWidth

            double colwidth;                                                                //default or base column width (in pixel)
            var width = ((XSSFSheet)Sheet).worksheet.sheetFormatPr.defaultColWidth;         //string length with padding
            if(width != 0.0) {
                colwidth = width * MDW;
            } else {
                var length = ((XSSFSheet)Sheet).worksheet.sheetFormatPr.baseColWidth;       //string length with out padding
                var fontwidth = Math.Truncate((length * MDW + 5) / MDW * 256) / 256;
                var tmp = 256 * fontwidth + Math.Truncate(128 / MDW);
                colwidth = Math.Truncate((tmp / 256) * MDW) + 3;                            // +3 ???
            }
            int _left = Math.Min(dx1, dx2);
            int _top = Math.Min(dy1, dy2);
            int _right = Math.Max(dx1, dx2);
            int _bottom = Math.Max(dy1, dy2);

            cell1.colOff = EMUtoMakerCol(Sheet, MDW, colwidth, _left, cell1);
            cell1.rowOff = EMUtoMakerRow(Sheet, _top, cell1);
            cell2.colOff = EMUtoMakerCol(Sheet, MDW, colwidth, _right, cell2);
            cell2.rowOff = EMUtoMakerRow(Sheet, _bottom, cell2);

            this.left   = Math.Min(dx1, dx2);
            this.top    = Math.Min(dy1, dy2);
            this.width  = Math.Abs(dx1 - dx2);
            this.height = Math.Abs(dy1 - dy2);
        }

        protected long EMUtoMakerCol(ISheet Sheet, double MDW, double Colwith, int EMU, CT_Marker Mkr) {
            double width_px;
            Mkr.colOff = EMU;
            Mkr.col = 0;
            for(int iCol = 0; iCol < SpreadsheetVersion.EXCEL2007.MaxColumns; iCol++) {
                width_px = Colwith;
                foreach(var cols in ((XSSFSheet)Sheet).worksheet.cols) {
                    foreach(var col in cols.col) {
                        if(col.min <= iCol + 1 && iCol + 1 <= col.max) {
                            width_px = col.width * MDW;
                            goto lblforbreak;
                        }
                    }
                }
lblforbreak:
                int EMUwidth = Units.PixelToEMU((int)Math.Round(width_px, 1));
                if(Mkr.colOff >= EMUwidth) {
                    Mkr.colOff -= EMUwidth;
                    Mkr.col++;
                } else {
                    return Mkr.colOff;
                }
            }
            return -1;
        }

        protected long EMUtoMakerRow(ISheet Sheet, int EMU, CT_Marker Mkr) {
            Mkr.rowOff = EMU;
            Mkr.row= 0;
            for(int iRow = 0; iRow < SpreadsheetVersion.EXCEL2007.MaxRows; iRow++) {
                double height = ((XSSFSheet) Sheet).DefaultRowHeightInPoints;
                var row = (XSSFRow)((XSSFSheet) Sheet).GetRow(iRow);
                if(row != null) {
                    height = row.HeightInPoints;
                }
                if(Mkr.rowOff >= Units.ToEMU(height)) {
                    Mkr.rowOff -= Units.ToEMU(height);
                    Mkr.row++;
                } else {
                    return Mkr.rowOff;
                }
            }
            return -1;
        }

        /**
         * Create XSSFClientAnchor from existing xml beans
         *
         * @param cell1 starting anchor point
         * @param cell2 ending anchor point
         */
        internal XSSFClientAnchor(CT_Marker cell1, CT_Marker cell2, int left, int top, int right, int bottom)
        {
            this.cell1 = cell1;
            this.cell2 = cell2;

            this.left   = left;
            this.top    = top;
            this.width  = Math.Abs(right- left);
            this.height = Math.Abs(bottom - top);
        }

        /// <summary>
        /// </summary>
        /// <param name="sheet">sheet</param>
        /// <param name="cell">starting point and offsets (may be zeros)</param>
        /// <param name="size">dimensions to calculate relative to starting point</param>
        private CT_Marker calcCell(CT_Marker cell, long w, long h)
        {
            CT_Marker c2 = new CT_Marker();

            int r = cell.row;
            int c = cell.col;

            int cw = Units.ColumnWidthToEMU((int)sheet.GetColumnWidth(c));

            // start with width - offset, then keep adding column widths until the next one Puts us over w
            long wPos = cw - cell.colOff;

            while (wPos < w)
            {
                c++;
                cw = Units.ColumnWidthToEMU((int)sheet.GetColumnWidth(c));
                wPos += cw;
            }
            // now wPos >= w, so end column = c, now figure offset
            c2.col = (c);
            c2.colOff = (cw - (wPos - w));

            int rh = Units.ToEMU(GetRowHeight(sheet, r));
            // start with height - offset, then keep adding row heights until the next one Puts us over h
            long hPos = rh - cell.rowOff;

            while (hPos < h)
            {
                r++;
                rh = Units.ToEMU(GetRowHeight(sheet, r));
                hPos += rh;
            }
            // now hPos >= h, so end row = r, now figure offset
            c2.row = (r);
            c2.rowOff = (rh - (hPos - h));

            return c2;
        }

        /// <summary>
        /// </summary>
        /// <param name="sheet">sheet</param>
        /// <param name="row">row</param>
        /// <return>height in twips (1/20th of point) for row or default</return>
        private static float GetRowHeight(XSSFSheet sheet, int row)
        {
            XSSFRow r = sheet.GetRow(row) as XSSFRow;
            return r == null ? sheet.DefaultRowHeightInPoints : r.HeightInPoints;
        }

        private CT_Marker Cell1
        {
            get
            {
                return cell1 != null ? cell1 : calcCell(EMPTY_MARKER, position.x, position.y);
            }
        }

        private CT_Marker Cell2
        {
            get
            {
                return cell2 != null ? cell2 : calcCell(Cell1, size.cx, size.cy);
            }
            
        }

        public override bool Equals(Object o)
        {
            if (o == null || o is not XSSFClientAnchor anchor) return false;

            return Dx1 == anchor.Dx1 &&
                   Dx2 == anchor.Dx2 &&
                   Dy1 == anchor.Dy1 &&
                   Dy2 == anchor.Dy2 &&
                   Col1 == anchor.Col1 &&
                   Col2 == anchor.Col2 &&
                   Row1 == anchor.Row1 &&
                   Row2 == anchor.Row2;

        }

        public override int GetHashCode()
        {
            return 42; // any arbitrary constant will do
        }

        public override String ToString()
        {
            return "from : " + Cell1.ToString() + "; to: " + Cell2.ToString();
        }

        /**
         * Return starting anchor point
         *
         * @return starting anchor point
         */

        internal CT_Marker From
        {
            get
            {
                return Cell1;
            }
            set 
            {
                cell1 = value;
            }
        }

        /**
         * Return ending anchor point
         *
         * @return ending anchor point
         */

        public CT_Marker To
        {
            get
            {
                return Cell2;
            }
            set 
            {
                cell2 = value;
            }
        }


         internal bool IsSet()
         {
            CT_Marker c1 = Cell1;
            CT_Marker c2 = Cell2;
            return !(c1.col == 0 && c2.col == 0 &&
                     c1.row == 0 && c2.row == 0);
        }

        /// <summary>absolute top-left position, or null if position is determined from the "from" cell
        /// To use this, "from" must be Set to null.
        /// </summary>
        /// <return></return>
        /// @since POI 3.17 beta 1
        public CT_Point2D Position
        {
            get
            {
                return position;
            }
            set { position = value; }
        }

        /// <summary>size or null, if size is determined from the to and from cells
        /// To use this, "to" must be Set to null.
        /// </summary>
        /// @since POI 3.17 beta 1
        public CT_PositiveSize2D Size
        {
            get { return size; }
            set { size = value; }
        }

        #region IClientAnchor Members
        public override int Dx1
        {
            get
            {
                return (int)Cell1.colOff;
            }
            set
            {
                cell1.colOff = value;
            }
        }

        public override int Dy1
        {
            get
            {
                return (int)Cell1.rowOff;
            }
            set
            {
                cell1.rowOff = value;
            }
        }

        public override int Dy2
        {
            get
            {
                return (int)Cell2.rowOff;
            }
            set
            {
                cell2.rowOff = value;
            }
        }

        public override int Dx2
        {
            get
            {
                return (int)Cell2.colOff;
            }
            set
            {
                cell2.colOff = value;
            }
        }
        public AnchorType AnchorType
        {
            get
            {
                return this.anchorType;
            }
            set
            {
                this.anchorType = value;
            }
        }

        public int Col1
        {
            get
            {
                return cell1.col;
            }
            set
            {
                Cell1.col=value;
            }
        }

        public int Col2
        {
            get
            {
                return Cell2.col;
            }
            set
            {
                cell2.col = value;
            }
        }

        public int Row1
        {
            get
            {
                return Cell1.row;
            }
            set
            {
                cell1.row = value;
            }
        }

        public int Row2
        {
            get
            {
                return Cell2.row;
            }
            set
            {
                cell2.row = value;
            }
        }

        #endregion
    }
}



