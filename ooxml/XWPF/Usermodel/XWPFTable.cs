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
namespace NPOI.XWPF.UserModel
{
    using System;
    using System.Text;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;

    /**
     * <p>Sketch of XWPFTable class. Only table's text is being hold.</p>
     * <p>Specifies the contents of a table present in the document. A table is a set
     * of paragraphs (and other block-level content) arranged in rows and columns.</p>
     */
    public class XWPFTable : IBodyElement, ISDTContents
    {

        protected StringBuilder text = new StringBuilder();
        private CT_Tbl ctTbl;
        protected List<XWPFTableRow> tableRows;
        protected List<String> styleIDs;

        // Create a map from this XWPF-level enum to the STBorder.Enum values
        public enum XWPFBorderType { NIL, NONE, SINGLE, THICK, DOUBLE, DOTTED, DASHED, DOT_DASH };
        internal static Dictionary<XWPFBorderType, ST_Border> xwpfBorderTypeMap;
        // Create a map from the STBorder.Enum values to the XWPF-level enums
        internal static Dictionary<ST_Border, XWPFBorderType> stBorderTypeMap;

        protected IBody part;
        static XWPFTable()
        {
            // populate enum maps
            xwpfBorderTypeMap = new Dictionary<XWPFBorderType, ST_Border>();
            xwpfBorderTypeMap.Add(XWPFBorderType.NIL, ST_Border.nil);
            xwpfBorderTypeMap.Add(XWPFBorderType.NONE, ST_Border.none);
            xwpfBorderTypeMap.Add(XWPFBorderType.SINGLE, ST_Border.single);
            xwpfBorderTypeMap.Add(XWPFBorderType.THICK, ST_Border.thick);
            xwpfBorderTypeMap.Add(XWPFBorderType.DOUBLE, ST_Border.@double);
            xwpfBorderTypeMap.Add(XWPFBorderType.DOTTED, ST_Border.dotted);
            xwpfBorderTypeMap.Add(XWPFBorderType.DASHED, ST_Border.dashed);
            xwpfBorderTypeMap.Add(XWPFBorderType.DOT_DASH, ST_Border.dotDash);

            stBorderTypeMap = new Dictionary<ST_Border, XWPFBorderType>();
            stBorderTypeMap.Add(ST_Border.nil, XWPFBorderType.NIL);
            stBorderTypeMap.Add(ST_Border.none, XWPFBorderType.NONE);
            stBorderTypeMap.Add(ST_Border.single, XWPFBorderType.SINGLE);
            stBorderTypeMap.Add(ST_Border.thick, XWPFBorderType.THICK);
            stBorderTypeMap.Add(ST_Border.@double, XWPFBorderType.DOUBLE);
            stBorderTypeMap.Add(ST_Border.dotted, XWPFBorderType.DOTTED);
            stBorderTypeMap.Add(ST_Border.dashed, XWPFBorderType.DASHED);
            stBorderTypeMap.Add(ST_Border.dotDash, XWPFBorderType.DOT_DASH);
        }
        public XWPFTable(CT_Tbl table, IBody part, int row, int col)
            : this(table, part)
        {

            CT_TblGrid ctTblGrid = table.AddNewTblGrid();
            for (int j = 0; j < col; j++)
            {
                CT_TblGridCol ctGridCol= ctTblGrid.AddNewGridCol();
                ctGridCol.w = 300;
            }
            for (int i = 0; i < row; i++)
            {
                XWPFTableRow tabRow = (GetRow(i) == null) ? CreateRow() : GetRow(i);
                for (int k = 0; k < col; k++)
                {
                    if (tabRow.GetCell(k) == null)
                    {
                        tabRow.CreateCell();
                    }
                }
            }
        }

        public void SetColumnWidth(int columnIndex, ulong width)
        {
            if (this.ctTbl.tblGrid == null)
                return;

            if (columnIndex > this.ctTbl.tblGrid.gridCol.Count)
            {
                throw new ArgumentOutOfRangeException(string.Format("Column index {0} doesn't exist.", columnIndex));
            }
            this.ctTbl.tblGrid.gridCol[columnIndex].w = width;

        }

        public XWPFTable(CT_Tbl table, IBody part)
        {
            this.part = part;
            this.ctTbl = table;

            tableRows = new List<XWPFTableRow>();
            // is an empty table: I add one row and one column as default
            if (table.SizeOfTrArray() == 0)
                CreateEmptyTable(table);

            foreach (CT_Row row in table.GetTrList()) {
                StringBuilder rowText = new StringBuilder();
                row.Table = table;
                XWPFTableRow tabRow = new XWPFTableRow(row, this);
                tableRows.Add(tabRow);
                foreach (CT_Tc cell in row.GetTcList()) {
                    foreach (CT_P ctp in cell.GetPList()) {
                        XWPFParagraph p = new XWPFParagraph(ctp, part);
                        if (rowText.Length > 0) {
                            rowText.Append('\t');
                        }
                        rowText.Append(p.Text);
                    }
                }
                if (rowText.Length > 0) {
                    this.text.Append(rowText);
                    this.text.Append('\n');
                }
            }
        }
        
        private void CreateEmptyTable(CT_Tbl table)
        {
            // MINIMUM ELEMENTS FOR A TABLE
            table.AddNewTr().AddNewTc().AddNewP();

            CT_TblPr tblpro = table.AddNewTblPr();
            if (!tblpro.IsSetTblW())
                tblpro.AddNewTblW().w = "0";
            tblpro.tblW.type=(ST_TblWidth.auto);

            // layout
             tblpro.AddNewTblLayout().type =  ST_TblLayoutType.autofit;

            // borders
            CT_TblBorders borders = tblpro.AddNewTblBorders();
            borders.AddNewBottom().val=ST_Border.single;
            borders.AddNewInsideH().val = ST_Border.single;
            borders.AddNewInsideV().val = ST_Border.single;
            borders.AddNewLeft().val = ST_Border.single;
            borders.AddNewRight().val = ST_Border.single;
            borders.AddNewTop().val = ST_Border.single;

            
            //CT_TblGrid tblgrid=table.AddNewTblGrid();
            //tblgrid.AddNewGridCol().w= (ulong)2000;
           
        }

        /**
         * @return ctTbl object
         */

        internal CT_Tbl GetCTTbl()
        {
            return ctTbl;
        }

        /**
         * Convenience method to extract text in cells.  This
         * does not extract text recursively in cells, and it does not
         * currently include text in SDT (form) components.
         * <p>
         * To get all text within a table, see XWPFWordExtractor's appendTableText
         * as an example. 
         *
         * @return text
         */
        public String Text
        {
            get
            {
                return text.ToString();
            }
        }

        public void AddNewRowBetween(int start, int end)
        {
            throw new NotImplementedException();
        }

        /**
         * add a new column for each row in this table
         */
        public void AddNewCol()
        {
            if (ctTbl.SizeOfTrArray() == 0) {
                CreateRow();
            }
            for (int i = 0; i < ctTbl.SizeOfTrArray(); i++) {
                XWPFTableRow tabRow = new XWPFTableRow(ctTbl.GetTrArray(i), this);
                tabRow.CreateCell();
            }
        }

        /**
         * create a new XWPFTableRow object with as many cells as the number of columns defined in that moment
         *
         * @return tableRow
         */
        public XWPFTableRow CreateRow()
        {
            int sizeCol = ctTbl.SizeOfTrArray() > 0 ? ctTbl.GetTrArray(0)
                    .SizeOfTcArray() : 0;
            XWPFTableRow tabRow = new XWPFTableRow(ctTbl.AddNewTr(), this);
            AddColumn(tabRow, sizeCol);
            tableRows.Add(tabRow);
            return tabRow;
        }

        /**
         * @param pos - index of the row
         * @return the row at the position specified or null if no rows is defined or if the position is greather than the max size of rows array
         */
        public XWPFTableRow GetRow(int pos)
        {
            if (pos >= 0 && pos < ctTbl.SizeOfTrArray()) {
                //return new XWPFTableRow(ctTbl.GetTrArray(pos));
                return Rows[(pos)];
            }
            return null;
        }


        /**
         * @return width value
         */
        public int Width
        {
            get
            {
                CT_TblPr tblPr = GetTrPr();
                return tblPr.IsSetTblW() ? int.Parse(tblPr.tblW.w) : -1;
            }
            set 
            {

                CT_TblPr tblPr = GetTrPr();
                CT_TblWidth tblWidth = tblPr.IsSetTblW() ? tblPr.tblW : tblPr
                        .AddNewTblW();
                tblWidth.w = value.ToString();
                tblWidth.type = ST_TblWidth.pct;
            }
        }

        /**
         * @return number of rows in table
         */
        public int NumberOfRows
        {
            get
            {
                return ctTbl.SizeOfTrArray();
            }
        }

        private CT_TblPr GetTrPr()
        {
            return (ctTbl.tblPr != null) ? ctTbl.tblPr : ctTbl
                    .AddNewTblPr();
        }

        private void AddColumn(XWPFTableRow tabRow, int sizeCol)
        {
            if (sizeCol > 0)
            {
                for (int i = 0; i < sizeCol; i++)
                {
                    tabRow.CreateCell();
                }
            }
        }

        /**
         * Get the StyleID of the table
         * @return	style-ID of the table
         */
        public String StyleID
        {
            get
            {
                String styleId = null;
                CT_TblPr tblPr = ctTbl.tblPr;
                if (tblPr != null)
                {
                    CT_String styleStr = tblPr.tblStyle;
                    if (styleStr != null)
                    {
                        styleId = styleStr.val;
                    }
                }
                return styleId;
            }
            set
            {
                CT_TblPr tblPr = GetTrPr();
                CT_String styleStr = tblPr.tblStyle;
                if (styleStr == null)
                {
                    styleStr = tblPr.AddNewTblStyle();
                }
                styleStr.val = value;
            }
        }
        public XWPFBorderType InsideHBorderType
        {
            get
            {
                XWPFBorderType bt = XWPFBorderType.NONE;

                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblBorders())
                {
                    CT_TblBorders ctb = tblPr.tblBorders;
                    if (ctb.IsSetInsideH())
                    {
                        CT_Border border = ctb.insideH;
                        bt = stBorderTypeMap[border.val];
                    }
                }
                return bt;
            }
        }

        public int InsideHBorderSize
        {
            get
            {
                int size = -1;

                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblBorders())
                {
                    CT_TblBorders ctb = tblPr.tblBorders;
                    if (ctb.IsSetInsideH())
                    {
                        CT_Border border = ctb.insideH;
                        size = (int)border.sz;
                    }
                }
                return size;
            }
        }

        public int InsideHBorderSpace
        {
            get
            {
                int space = -1;

                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblBorders())
                {
                    CT_TblBorders ctb = tblPr.tblBorders;
                    if (ctb.IsSetInsideH())
                    {
                        CT_Border border = ctb.insideH;
                        space = (int)border.space;
                    }
                }
                return space;
            }
        }

        public String InsideHBorderColor
        {
            get
            {
                String color = null;

                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblBorders())
                {
                    CT_TblBorders ctb = tblPr.tblBorders;
                    if (ctb.IsSetInsideH())
                    {
                        CT_Border border = ctb.insideH;
                        color = border.color;
                    }
                }
                return color;
            }
        }

        public XWPFBorderType InsideVBorderType
        {
            get
            {
                XWPFBorderType bt = XWPFBorderType.NONE;

                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblBorders())
                {
                    CT_TblBorders ctb = tblPr.tblBorders;
                    if (ctb.IsSetInsideV())
                    {
                        CT_Border border = ctb.insideV;
                        bt = stBorderTypeMap[border.val];
                    }
                }

                return bt;
            }
        }

        public int InsideVBorderSize
        {
            get
            {
                int size = -1;

                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblBorders())
                {
                    CT_TblBorders ctb = tblPr.tblBorders;
                    if (ctb.IsSetInsideV())
                    {
                        CT_Border border = ctb.insideV;
                        size = (int)border.sz;
                    }
                }
                return size;
            }
        }

        public int InsideVBorderSpace
        {
            get
            {
                int space = -1;

                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblBorders())
                {
                    CT_TblBorders ctb = tblPr.tblBorders;
                    if (ctb.IsSetInsideV())
                    {
                        CT_Border border = ctb.insideV;
                        space = (int)border.space;
                    }
                }
                return space;
            }
        }

        public String InsideVBorderColor
        {
            get
            {
                String color = null;

                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblBorders())
                {
                    CT_TblBorders ctb = tblPr.tblBorders;
                    if (ctb.IsSetInsideV())
                    {
                        CT_Border border = ctb.insideV;
                        color = border.color;
                    }
                }
                return color;
            }
        }

        public int RowBandSize
        {
            get
            {
                int size = 0;
                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblStyleRowBandSize())
                {
                    CT_DecimalNumber rowSize = tblPr.tblStyleRowBandSize;
                    int.TryParse(rowSize.val, out size);
                }
                return size;
            }
            set 
            {
                CT_TblPr tblPr = GetTrPr();
                CT_DecimalNumber rowSize = tblPr.IsSetTblStyleRowBandSize() ? tblPr.tblStyleRowBandSize : tblPr.AddNewTblStyleRowBandSize();
                rowSize.val = value.ToString();			
            }
        }

        public int ColBandSize
        {
            get
            {
                int size = 0;
                CT_TblPr tblPr = GetTrPr();
                if (tblPr.IsSetTblStyleColBandSize())
                {
                    CT_DecimalNumber colSize = tblPr.tblStyleColBandSize;
                    int.TryParse(colSize.val, out size);
                }
                return size;
            }
            set 
            {
                CT_TblPr tblPr = GetTrPr();
                CT_DecimalNumber colSize = tblPr.IsSetTblStyleColBandSize() ? tblPr.tblStyleColBandSize : tblPr.AddNewTblStyleColBandSize();
                colSize.val = value.ToString();
            }
        }
        public void SetTopBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_TblPr tblPr = GetTrPr();
            CT_TblBorders ctb = tblPr.IsSetTblBorders() ? tblPr.tblBorders : tblPr.AddNewTblBorders();
            CT_Border b = ctb.top!=null ? ctb.top : ctb.AddNewTop();
            b.val = xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }
        public void SetBottomBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_TblPr tblPr = GetTrPr();
            CT_TblBorders ctb = tblPr.IsSetTblBorders() ? tblPr.tblBorders : tblPr.AddNewTblBorders();
            CT_Border b = ctb.bottom != null ? ctb.bottom : ctb.AddNewBottom();
            b.val = xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }
        public void SetLeftBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_TblPr tblPr = GetTrPr();
            CT_TblBorders ctb = tblPr.IsSetTblBorders() ? tblPr.tblBorders : tblPr.AddNewTblBorders();
            CT_Border b = ctb.left != null ? ctb.left : ctb.AddNewLeft();
            b.val = xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }
        public void SetRightBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_TblPr tblPr = GetTrPr();
            CT_TblBorders ctb = tblPr.IsSetTblBorders() ? tblPr.tblBorders : tblPr.AddNewTblBorders();
            CT_Border b = ctb.right != null ? ctb.right : ctb.AddNewRight();
            b.val = xwpfBorderTypeMap[type];
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }
        public void SetInsideHBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_TblPr tblPr = GetTrPr();
            CT_TblBorders ctb = tblPr.IsSetTblBorders() ? tblPr.tblBorders : tblPr.AddNewTblBorders();
            CT_Border b = ctb.IsSetInsideH() ? ctb.insideH : ctb.AddNewInsideH();
            b.val = (xwpfBorderTypeMap[(type)]);
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }

        public void SetInsideVBorder(XWPFBorderType type, int size, int space, String rgbColor)
        {
            CT_TblPr tblPr = GetTrPr();
            CT_TblBorders ctb = tblPr.IsSetTblBorders() ? tblPr.tblBorders : tblPr.AddNewTblBorders();
            CT_Border b = ctb.IsSetInsideV() ? ctb.insideV : ctb.AddNewInsideV();
            b.val = (xwpfBorderTypeMap[type]);
            b.sz = (ulong)size;
            b.space = (ulong)space;
            b.color = (rgbColor);
        }

        public int CellMarginTop
        {
            get
            {
                int margin = 0;
                CT_TblPr tblPr = GetTrPr();
                CT_TblCellMar tcm = tblPr.tblCellMar;
                if (tcm != null)
                {
                    CT_TblWidth tw = tcm.top;
                    if (tw != null)
                    {
                        int.TryParse(tw.w, out margin);
                    }
                }
                return margin;
            }
        }

        public int CellMarginLeft
        {
            get
            {
                int margin = 0;
                CT_TblPr tblPr = GetTrPr();
                CT_TblCellMar tcm = tblPr.tblCellMar;
                if (tcm != null)
                {
                    CT_TblWidth tw = tcm.left;
                    if (tw != null)
                    {
                        int.TryParse(tw.w, out margin);
                    }
                }
                return margin;
            }
        }

        public int CellMarginBottom
        {
            get
            {
                int margin = 0;
                CT_TblPr tblPr = GetTrPr();
                CT_TblCellMar tcm = tblPr.tblCellMar;
                if (tcm != null)
                {
                    CT_TblWidth tw = tcm.bottom;
                    if (tw != null)
                    {
                        int.TryParse(tw.w, out margin);
                    }
                }
                return margin;
            }
        }

        public int CellMarginRight
        {
            get
            {
                int margin = 0;
                CT_TblPr tblPr = GetTrPr();
                CT_TblCellMar tcm = tblPr.tblCellMar;
                if (tcm != null)
                {
                    CT_TblWidth tw = tcm.right;
                    if (tw != null)
                    {
                        int.TryParse(tw.w, out margin);
                    }
                }
                return margin;
            }
        }

        public void SetCellMargins(int top, int left, int bottom, int right)
        {
            CT_TblPr tblPr = GetTrPr();
            CT_TblCellMar tcm = tblPr.IsSetTblCellMar() ? tblPr.tblCellMar : tblPr.AddNewTblCellMar();

            CT_TblWidth tw = tcm.IsSetLeft() ? tcm.left : tcm.AddNewLeft();
            tw.type = (ST_TblWidth.dxa);
            tw.w = left.ToString();

            tw = tcm.IsSetTop() ? tcm.top : tcm.AddNewTop();
            tw.type = (ST_TblWidth.dxa);
            tw.w = top.ToString();

            tw = tcm.IsSetBottom() ? tcm.bottom : tcm.AddNewBottom();
            tw.type = (ST_TblWidth.dxa);
            tw.w = bottom.ToString();

            tw = tcm.IsSetRight() ? tcm.right : tcm.AddNewRight();
            tw.type = (ST_TblWidth.dxa);
            tw.w = right.ToString();
        }
    
        /**
         * add a new Row to the table
         * 
         * @param row	the row which should be Added
         */
        public void AddRow(XWPFTableRow row)
        {
            ctTbl.AddNewTr();
            ctTbl.SetTrArray(this.NumberOfRows-1, row.GetCTRow());
            tableRows.Add(row);
        }

        /**
         * add a new Row to the table
         * at position pos
         * @param row	the row which should be Added
         */
        public bool AddRow(XWPFTableRow row, int pos)
        {
            if (pos >= 0 && pos <= tableRows.Count)
            {
                ctTbl.InsertNewTr(pos);
                ctTbl.SetTrArray(pos, row.GetCTRow());
                tableRows.Insert(pos, row);
                return true;
            }
            return false;
        }

        /**
         * inserts a new tablerow 
         * @param pos
         * @return  the inserted row
         */
        public XWPFTableRow InsertNewTableRow(int pos)
        {
            if(pos >= 0 && pos <= tableRows.Count){
                CT_Row row = ctTbl.InsertNewTr(pos);
                XWPFTableRow tableRow = new XWPFTableRow(row, this);
                tableRows.Insert(pos, tableRow);
                return tableRow;
            }
            return null;
        }


        /**
         * Remove a row at position pos from the table
         * @param pos	position the Row in the Table
         */
        public bool RemoveRow(int pos)
        {
            if (pos >= 0 && pos < tableRows.Count) {
                if (ctTbl.SizeOfTrArray() > 0)
                {
                    ctTbl.RemoveTr(pos);
                }
                tableRows.RemoveAt(pos);
                return true;
            }
            return false;
        }

        public List<XWPFTableRow> Rows
        {
            get
            {
                return tableRows;
            }
        }


        /**
         * returns the type of the BodyElement Table
         * @see NPOI.XWPF.UserModel.IBodyElement#getElementType()
         */
        public BodyElementType ElementType
        {
            get
            {
                return BodyElementType.TABLE;
            }
        }

        public IBody Body
        {
            get
            {
                return part;
            }
        }

        /**
         * returns the part of the bodyElement
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
        public POIXMLDocumentPart Part
        {
            get
            {
                if (part != null)
                {
                    return part.Part;
                }
                return null;
            }
        }

        /**
         * returns the partType of the bodyPart which owns the bodyElement
         * @see NPOI.XWPF.UserModel.IBody#getPartType()
         */
        public BodyType PartType
        {
            get
            {
                return part.PartType;
            }
        }

        /**
         * returns the XWPFRow which belongs to the CTRow row
         * if this row is not existing in the table null will be returned
         */
        public XWPFTableRow GetRow(CT_Row row)
        {
            for(int i=0; i<Rows.Count; i++){
                if(Rows[(i)].GetCTRow() == row) return GetRow(i); 
            }
            return null;
        }
    }// end class

}