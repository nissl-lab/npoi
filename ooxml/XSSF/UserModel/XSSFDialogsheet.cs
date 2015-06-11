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

using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
namespace NPOI.XSSF.UserModel
{

    //YK: TODO: this is only a prototype
    public class XSSFDialogsheet : XSSFSheet, ISheet
    {
        protected CT_Dialogsheet dialogsheet;

        public XSSFDialogsheet(XSSFSheet sheet)
            : base(sheet.GetPackagePart(), sheet.GetPackageRelationship())
        {

            this.dialogsheet = new CT_Dialogsheet();
            this.worksheet = new CT_Worksheet();
        }

        public override IRow CreateRow(int rowNum)
        {
            return null;
        }

        protected CT_HeaderFooter GetSheetTypeHeaderFooter()
        {
            if (dialogsheet.headerFooter == null)
            {
                dialogsheet.headerFooter = (new CT_HeaderFooter());
            }
            return dialogsheet.headerFooter;
        }

        protected CT_SheetPr GetSheetTypeSheetPr()
        {
            if (dialogsheet.sheetPr == null)
            {
                dialogsheet.sheetPr = (new CT_SheetPr());
            }
            return dialogsheet.sheetPr;
        }

        protected CT_PageBreak GetSheetTypeColumnBreaks()
        {
            return null;
        }

        protected CT_SheetFormatPr GetSheetTypeSheetFormatPr()
        {
            if (dialogsheet.sheetFormatPr == null)
            {
                dialogsheet.sheetFormatPr = (new CT_SheetFormatPr());
            }
            return dialogsheet.sheetFormatPr;
        }

        protected CT_PageMargins GetSheetTypePageMargins()
        {
            if (dialogsheet.pageMargins == null)
            {
                dialogsheet.pageMargins = (new CT_PageMargins());
            }
            return dialogsheet.pageMargins;
        }

        protected CT_PageBreak GetSheetTypeRowBreaks()
        {
            return null;
        }

        protected CT_SheetViews GetSheetTypeSheetViews()
        {
            if (dialogsheet.sheetViews == null)
            {
                dialogsheet.sheetViews = (new CT_SheetViews());
                dialogsheet.sheetViews.AddNewSheetView();
            }
            return dialogsheet.sheetViews;
        }

        protected CT_PrintOptions GetSheetTypePrintOptions()
        {
            if (dialogsheet.printOptions == null)
            {
                dialogsheet.printOptions = (new CT_PrintOptions());
            }
            return dialogsheet.printOptions;
        }

        protected CT_SheetProtection GetSheetTypeProtection()
        {
            if (dialogsheet.sheetProtection == null)
            {
                dialogsheet.sheetProtection = (new CT_SheetProtection());
            }
            return dialogsheet.sheetProtection;
        }

        public bool GetDialog()
        {
            return true;
        }

        IRow ISheet.CreateRow(int rownum)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.RemoveRow(IRow row)
        {
            throw new System.NotImplementedException();
        }

        IRow ISheet.GetRow(int rownum)
        {
            throw new System.NotImplementedException();
        }

        int ISheet.PhysicalNumberOfRows
        {
            get { throw new System.NotImplementedException(); }
        }

        int ISheet.FirstRowNum
        {
            get { throw new System.NotImplementedException(); }
        }

        int ISheet.LastRowNum
        {
            get { throw new System.NotImplementedException(); }
        }

        bool ISheet.ForceFormulaRecalculation
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        void ISheet.SetColumnHidden(int columnIndex, bool hidden)
        {
            throw new System.NotImplementedException();
        }

        bool ISheet.IsColumnHidden(int columnIndex)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetColumnWidth(int columnIndex, int width)
        {
            throw new System.NotImplementedException();
        }

        int ISheet.GetColumnWidth(int columnIndex)
        {
            throw new System.NotImplementedException();
        }

        int ISheet.DefaultColumnWidth
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        short ISheet.DefaultRowHeight
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        float ISheet.DefaultRowHeightInPoints
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        ICellStyle ISheet.GetColumnStyle(int column)
        {
            throw new System.NotImplementedException();
        }

        int ISheet.AddMergedRegion(SS.Util.CellRangeAddress region)
        {
            throw new System.NotImplementedException();
        }

        bool ISheet.HorizontallyCenter
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.VerticallyCenter
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        void ISheet.RemoveMergedRegion(int index)
        {
            throw new System.NotImplementedException();
        }

        int ISheet.NumMergedRegions
        {
            get { throw new System.NotImplementedException(); }
        }

        SS.Util.CellRangeAddress ISheet.GetMergedRegion(int index)
        {
            throw new System.NotImplementedException();
        }

        System.Collections.IEnumerator ISheet.GetRowEnumerator()
        {
            throw new System.NotImplementedException();
        }

        bool ISheet.DisplayZeros
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.Autobreaks
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.DisplayGuts
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.FitToPage
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.RowSumsBelow
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.RowSumsRight
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.IsPrintGridlines
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        IPrintSetup ISheet.PrintSetup
        {
            get { throw new System.NotImplementedException(); }
        }

        IHeader ISheet.Header
        {
            get { throw new System.NotImplementedException(); }
        }

        IFooter ISheet.Footer
        {
            get { throw new System.NotImplementedException(); }
        }

        double ISheet.GetMargin(MarginType margin)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetMargin(MarginType margin, double size)
        {
            throw new System.NotImplementedException();
        }

        bool ISheet.Protect
        {
            get { throw new System.NotImplementedException(); }
        }

        void ISheet.ProtectSheet(string password)
        {
            throw new System.NotImplementedException();
        }

        bool ISheet.ScenarioProtect
        {
            get { throw new System.NotImplementedException(); }
        }

        short ISheet.TabColorIndex
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        IDrawing ISheet.DrawingPatriarch
        {
            get { throw new System.NotImplementedException(); }
        }

        void ISheet.SetZoom(int numerator, int denominator)
        {
            throw new System.NotImplementedException();
        }

        short ISheet.TopRow
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        short ISheet.LeftCol
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        void ISheet.ShowInPane(short toprow, short leftcol)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.ShiftRows(int startRow, int endRow, int n)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.CreateFreezePane(int colSplit, int rowSplit)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane)
        {
            throw new System.NotImplementedException();
        }

        SS.Util.PaneInformation ISheet.PaneInformation
        {
            get { throw new System.NotImplementedException(); }
        }

        bool ISheet.DisplayGridlines
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.DisplayFormulas
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.DisplayRowColHeadings
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.IsActive
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        bool ISheet.IsRowBroken(int row)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.RemoveRowBreak(int row)
        {
            throw new System.NotImplementedException();
        }

        int[] ISheet.RowBreaks
        {
            get { throw new System.NotImplementedException(); }
        }

        int[] ISheet.ColumnBreaks
        {
            get { throw new System.NotImplementedException(); }
        }

        void ISheet.SetActiveCell(int row, int column)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetActiveCellRange(System.Collections.Generic.List<SS.Util.CellRangeAddress8Bit> cellranges, int activeRange, int activeRow, int activeColumn)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetColumnBreak(int column)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetRowBreak(int row)
        {
            throw new System.NotImplementedException();
        }

        bool ISheet.IsColumnBroken(int column)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.RemoveColumnBreak(int column)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.GroupColumn(int fromColumn, int toColumn)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.UngroupColumn(int fromColumn, int toColumn)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.GroupRow(int fromRow, int toRow)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.UngroupRow(int fromRow, int toRow)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetRowGroupCollapsed(int row, bool collapse)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.SetDefaultColumnStyle(int column, ICellStyle style)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.AutoSizeColumn(int column)
        {
            throw new System.NotImplementedException();
        }

        void ISheet.AutoSizeColumn(int column, bool useMergedCells)
        {
            throw new System.NotImplementedException();
        }

        IComment ISheet.GetCellComment(int row, int column)
        {
            throw new System.NotImplementedException();
        }

        IDrawing ISheet.CreateDrawingPatriarch()
        {
            throw new System.NotImplementedException();
        }

        IWorkbook ISheet.Workbook
        {
            get { throw new System.NotImplementedException(); }
        }

        string ISheet.SheetName
        {
            get { throw new System.NotImplementedException(); }
        }

        bool ISheet.IsSelected
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        void ISheet.SetActive(bool sel)
        {
            throw new System.NotImplementedException();
        }

        ICellRange<ICell> ISheet.SetArrayFormula(string formula, SS.Util.CellRangeAddress range)
        {
            throw new System.NotImplementedException();
        }

        ICellRange<ICell> ISheet.RemoveArrayFormula(ICell cell)
        {
            throw new System.NotImplementedException();
        }

        bool ISheet.IsMergedRegion(SS.Util.CellRangeAddress mergedRegion)
        {
            throw new System.NotImplementedException();
        }

        IDataValidationHelper ISheet.GetDataValidationHelper()
        {
            throw new System.NotImplementedException();
        }

        void ISheet.AddValidationData(IDataValidation dataValidation)
        {
            throw new System.NotImplementedException();
        }

        IAutoFilter ISheet.SetAutoFilter(SS.Util.CellRangeAddress range)
        {
            throw new System.NotImplementedException();
        }

        ISheetConditionalFormatting ISheet.SheetConditionalFormatting
        {
            get { throw new System.NotImplementedException(); }
        }

        bool IsRightToLeft
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
                throw new System.NotImplementedException();
            }
        }
    }
}


