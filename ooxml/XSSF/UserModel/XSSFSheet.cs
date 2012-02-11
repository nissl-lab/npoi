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
using NPOI.Util;
using System.IO;
using System.Xml;
using NPOI.XSSF.Model;
using System.Collections.Generic;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using NPOI.SS.Util;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS;
using NPOI.XSSF.UserModel.Helpers;
using NPOI.HSSF.Record;
namespace NPOI.XSSF.UserModel
{
    /**
     * High level representation of a SpreadsheetML worksheet.
     *
     * <p>
     * Sheets are the central structures within a workbook, and are where a user does most of his spreadsheet work.
     * The most common type of sheet is the worksheet, which is represented as a grid of cells. Worksheet cells can
     * contain text, numbers, dates, and formulas. Cells can also be formatted.
     * </p>
     */
    public class XSSFSheet : POIXMLDocumentPart, ISheet
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(XSSFSheet));

        //TODO make the two variable below private!
        protected CT_Sheet sheet;
        protected CT_Worksheet worksheet;

        private Dictionary<int, XSSFRow> _rows;
        private List<XSSFHyperlink> hyperlinks;
        private ColumnHelper columnHelper;
        private CommentsTable sheetComments;
        /**
         * cache of master shared formulas in this sheet.
         * Master shared formula is the first formula in a group of shared formulas is saved in the f element.
         */
        private Dictionary<int, CT_CellFormula> sharedFormulas;
        private Dictionary<String, XSSFTable> tables;
        private List<CellRangeAddress> arrayFormulas;
        private XSSFDataValidationHelper dataValidationHelper;

        /**
         * Creates new XSSFSheet   - called by XSSFWorkbook to create a sheet from scratch.
         *
         * @see NPOI.XSSF.usermodel.XSSFWorkbook#CreateSheet()
         */
        protected XSSFSheet()
            : base()
        {

            dataValidationHelper = new XSSFDataValidationHelper(this);
            OnDocumentCreate();
        }

        /**
         * Creates an XSSFSheet representing the given namespace part and relationship.
         * Should only be called by XSSFWorkbook when Reading in an exisiting file.
         *
         * @param part - The namespace part that holds xml data represenring this sheet.
         * @param rel - the relationship of the given namespace part in the underlying OPC namespace
         */
        protected XSSFSheet(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

            dataValidationHelper = new XSSFDataValidationHelper(this);
        }

        /**
         * Returns the parent XSSFWorkbook
         *
         * @return the parent XSSFWorkbook
         */
        public XSSFWorkbook GetWorkbook()
        {
            return (XSSFWorkbook)GetParent();
        }

        /**
         * Initialize worksheet data when Reading in an exisiting file.
         */

        protected void OnDocumentRead()
        {
            try
            {
                Read(GetPackagePart().GetInputStream());
            }
            catch (IOException e)
            {
                throw new POIXMLException(e);
            }
        }

        protected void Read(Stream is1)
        {
            try
            {
                worksheet = WorksheetDocument.Factory.Parse(is1).getWorksheet();
            }
            catch (XmlException e)
            {
                throw new POIXMLException(e);
            }

            InitRows(worksheet);
            columnHelper = new ColumnHelper(worksheet);

            // Look for bits we're interested in
            foreach (POIXMLDocumentPart p in GetRelations())
            {
                if (p is CommentsTable)
                {
                    sheetComments = (CommentsTable)p;
                    break;
                }
                if (p is XSSFTable)
                {
                    tables[p.GetPackageRelationship().Id]=(XSSFTable)p;
                }
            }

            // Process external hyperlinks for the sheet, if there are any
            InitHyperlinks();
        }

        /**
         * Initialize worksheet data when creating a new sheet.
         */

        protected void onDocumentCreate()
        {
            worksheet = newSheet();
            InitRows(worksheet);
            columnHelper = new ColumnHelper(worksheet);
            hyperlinks = new List<XSSFHyperlink>();
        }

        private void InitRows(CT_Worksheet worksheet)
        {
            _rows = new Dictionary<int, XSSFRow>();
            tables = new Dictionary<String, XSSFTable>();
            sharedFormulas = new Dictionary<int, CT_CellFormula>();
            arrayFormulas = new List<CellRangeAddress>();
            foreach (CT_Row row in worksheet.sheetData)
            {
                XSSFRow r = new XSSFRow(row, this);
                _rows[r.GetRowNum()] =  r;
            }
        }

        /**
         * Read hyperlink relations, link them with CT_Hyperlink beans in this worksheet
         * and Initialize the internal array of XSSFHyperlink objects
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        private void InitHyperlinks()
        {
            hyperlinks = new List<XSSFHyperlink>();

            if (!worksheet.IsSetHyperlinks()) return;

            try
            {
                PackageRelationshipCollection hyperRels =
                    GetPackagePart().GetRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.Relation);

                // Turn each one into a XSSFHyperlink
                foreach (CT_Hyperlink hyperlink in worksheet.hyperlinks)
                {
                    PackageRelationship hyperRel = null;
                    if (hyperlink.id != null)
                    {
                        hyperRel = hyperRels.GetRelationshipByID(hyperlink.id);
                    }

                    hyperlinks.Add(new XSSFHyperlink(hyperlink, hyperRel));
                }
            }
            catch (InvalidFormatException e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * Create a new CT_Worksheet instance with all values set to defaults
         *
         * @return a new instance
         */
        private static CT_Worksheet newSheet()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_SheetFormatPr ctFormat = worksheet.AddNewSheetFormatPr();
            ctFormat.defaultRowHeight = (15.0);

            CT_SheetView ctView = worksheet.AddNewSheetViews().addNewSheetView();
            ctView.workbookViewId = (0);

            worksheet.AddNewDimension().@ref = "A1";

            worksheet.AddNewSheetData();

            CT_PageMargins ctMargins = worksheet.AddNewPageMargins();
            ctMargins.bottom = (0.75);
            ctMargins.footer = (0.3);
            ctMargins.header = (0.3);
            ctMargins.left = (0.7);
            ctMargins.right = (0.7);
            ctMargins.top = (0.75);

            return worksheet;
        }

        /**
         * Provide access to the CT_Worksheet bean holding this sheet's data
         *
         * @return the CT_Worksheet bean holding this sheet's data
         */

        public CT_Worksheet GetCTWorksheet()
        {
            return this.worksheet;
        }

        public ColumnHelper GetColumnHelper()
        {
            return columnHelper;
        }

        /**
         * Returns the name of this sheet
         *
         * @return the name of this sheet
         */
        public String GetSheetName()
        {
            return sheet.name;
        }

        /**
         * Adds a merged region of cells (hence those cells form one).
         *
         * @param region (rowfrom/colfrom-rowto/colto) to merge
         * @return index of this region
         */
        public int AddMergedRegion(CellRangeAddress region)
        {
            region.Validate(SpreadsheetVersion.EXCEL2007);

            // throw InvalidOperationException if the argument CellRangeAddress intersects with
            // a multi-cell array formula defined in this sheet
            validateArrayFormulas(region);

            CT_MergeCells ctMergeCells = worksheet.IsSetMergeCells() ? worksheet.mergeCells : worksheet.AddNewMergeCells();
            CT_MergeCell ctMergeCell = ctMergeCells.AddNewMergeCell();
            ctMergeCell.@ref = (region.FormatAsString());
            return ctMergeCells.sizeOfMergeCellArray();
        }

        private void validateArrayFormulas(CellRangeAddress region)
        {
            int firstRow = region.FirstRow;
            int firstColumn = region.FirstColumn;
            int lastRow = region.LastRow;
            int lastColumn = region.LastColumn;
            for (int rowIn = firstRow; rowIn <= lastRow; rowIn++)   
            {
                for (int colIn = firstColumn; colIn <= lastColumn; colIn++)
                {
                    XSSFRow row = GetRow(rowIn);
                    if (row == null) continue;

                    XSSFCell cell = row.GetCell(colIn);
                    if (cell == null) continue;

                    if (cell.IsPartOfArrayFormulaGroup())
                    {
                        CellRangeAddress arrayRange = cell.GetArrayFormulaRange();
                        if (arrayRange.NumberOfCells > 1 &&
                                (arrayRange.IsInRange(region.FirstRow, region.FirstColumn) ||
                                  arrayRange.IsInRange(region.FirstRow, region.FirstColumn)))
                        {
                            String msg = "The range " + region.FormatAsString() + " intersects with a multi-cell array formula. " +
                                    "You cannot merge cells of an array.";
                            throw new InvalidOperationException(msg);
                        }
                    }
                }
            }

        }

        /**
         * Adjusts the column width to fit the contents.
         *
         * This process can be relatively slow on large sheets, so this should
         *  normally only be called once per column, at the end of your
         *  Processing.
         *
         * @param column the column index
         */
        public void AutoSizeColumn(int column)
        {
            AutoSizeColumn(column, false);
        }

        /**
         * Adjusts the column width to fit the contents.
         * <p>
         * This process can be relatively slow on large sheets, so this should
         *  normally only be called once per column, at the end of your
         *  Processing.
         * </p>
         * You can specify whether the content of merged cells should be considered or ignored.
         *  Default is to ignore merged cells.
         *
         * @param column the column index
         * @param useMergedCells whether to use the contents of merged cells when calculating the width of the column
         */
        public void AutoSizeColumn(int column, bool useMergedCells)
        {
            double width = SheetUtil.GetColumnWidth(this, column, useMergedCells);

            if (width != -1)
            {
                width *= 256;
                int maxColumnWidth = 255 * 256; // The maximum column width for an individual cell is 255 characters
                if (width > maxColumnWidth)
                {
                    width = maxColumnWidth;
                }
                SetColumnWidth(column, (int)(width));
                columnHelper.SetColBestFit(column, true);
            }
        }

        /**
         * Create a new SpreadsheetML drawing. If this sheet already Contains a drawing - return that.
         *
         * @return a SpreadsheetML drawing
         */
        public XSSFDrawing CreateDrawingPatriarch()
        {
            XSSFDrawing drawing = null;
            CT_Drawing ctDrawing = GetCT_Drawing();
            if (ctDrawing == null)
            {
                //drawingNumber = #drawings.Count + 1
                int drawingNumber = GetPackagePart().Package.GetPartsByContentType(XSSFRelation.DRAWINGS.ContentType).Count + 1;
                drawing = (XSSFDrawing)CreateRelationship(XSSFRelation.DRAWINGS, XSSFFactory.GetInstance(), drawingNumber);
                String relId = drawing.GetPackageRelationship().Id;

                //add CT_Drawing element which indicates that this sheet Contains drawing components built on the drawingML platform.
                //The relationship Id references the part Containing the drawingML defInitions.
                ctDrawing = worksheet.AddNewDrawing();
                ctDrawing.id = (relId);
            }
            else
            {
                //search the referenced drawing in the list of the sheet's relations
                foreach (POIXMLDocumentPart p in GetRelations())
                {
                    if (p is XSSFDrawing)
                    {
                        XSSFDrawing dr = (XSSFDrawing)p;
                        String drId = dr.GetPackageRelationship().Id;
                        if (drId.Equals(ctDrawing.id))
                        {
                            drawing = dr;
                            break;
                        }
                        break;
                    }
                }
                if (drawing == null)
                {
                    logger.Log(POILogger.ERROR, "Can't find drawing with id=" + ctDrawing.id + " in the list of the sheet's relationships");
                }
            }
            return drawing;
        }

        /**
         * Get VML drawing for this sheet (aka 'legacy' drawig)
         *
         * @param autoCreate if true, then a new VML drawing part is Created
         *
         * @return the VML drawing of <code>null</code> if the drawing was not found and autoCreate=false
         */
        protected XSSFVMLDrawing GetVMLDrawing(bool autoCreate)
        {
            XSSFVMLDrawing drawing = null;
            CT_LegacyDrawing ctDrawing = GetCT_LegacyDrawing();
            if (ctDrawing == null)
            {
                if (autoCreate)
                {
                    //drawingNumber = #drawings.Count + 1
                    int drawingNumber = GetPackagePart().getPackage().getPartsByContentType(XSSFRelation.VML_DRAWINGS.getContentType()).Count + 1;
                    drawing = (XSSFVMLDrawing)CreateRelationship(XSSFRelation.VML_DRAWINGS, XSSFFactory.GetInstance(), drawingNumber);
                    String relId = drawing.GetPackageRelationship().id;

                    //add CT_LegacyDrawing element which indicates that this sheet Contains drawing components built on the drawingML platform.
                    //The relationship Id references the part Containing the drawing defInitions.
                    ctDrawing = worksheet.AddNewLegacyDrawing();
                    ctDrawing.id = relId;
                }
            }
            else
            {
                //search the referenced drawing in the list of the sheet's relations
                foreach (POIXMLDocumentPart p in GetRelations())
                {
                    if (p is XSSFVMLDrawing)
                    {
                        XSSFVMLDrawing dr = (XSSFVMLDrawing)p;
                        String drId = dr.GetPackageRelationship().id;
                        if (drId.Equals(ctDrawing.id))
                        {
                            drawing = dr;
                            break;
                        }
                        break;
                    }
                }
                if (drawing == null)
                {
                    logger.Log(POILogger.ERROR, "Can't find VML drawing with id=" + ctDrawing.id + " in the list of the sheet's relationships");
                }
            }
            return drawing;
        }

        protected CT_Drawing GetCT_Drawing()
        {
            return worksheet.drawing;
        }
        protected CT_LegacyDrawing GetCT_LegacyDrawing()
        {
            return worksheet.legacyDrawing;
        }

        /**
         * Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
         * @param colSplit      Horizonatal position of split.
         * @param rowSplit      Vertical position of split.
         */
        public void CreateFreezePane(int colSplit, int rowSplit)
        {
            CreateFreezePane(colSplit, rowSplit, colSplit, rowSplit);
        }

        /**
         * Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
         *
         * <p>
         *     If both colSplit and rowSplit are zero then the existing freeze pane is Removed
         * </p>
         *
         * @param colSplit      Horizonatal position of split.
         * @param rowSplit      Vertical position of split.
         * @param leftmostColumn   Left column visible in right pane.
         * @param topRow        Top row visible in bottom pane
         */
        public void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
        {
            CT_SheetView ctView = GetDefaultSheetView();

            // If both colSplit and rowSplit are zero then the existing freeze pane is Removed
            if (colSplit == 0 && rowSplit == 0)
            {
                if (ctView.IsSetPane()) ctView.unSetPane();
                ctView.SetSelectionArray(null);
                return;
            }

            if (!ctView.IsSetPane())
            {
                ctView.AddNewPane();
            }
            CT_Pane pane = ctView.pane;

            if (colSplit > 0)
            {
                pane.SetXSplit(colSplit);
            }
            else
            {
                if (pane.IsSetXSplit()) pane.unSetXSplit();
            }
            if (rowSplit > 0)
            {
                pane.SetYSplit(rowSplit);
            }
            else
            {
                if (pane.IsSetYSplit()) pane.unSetYSplit();
            }

            pane.state = (ST_PaneState.frozen);
            if (rowSplit == 0)
            {
                pane.topLeftCell = (new CellReference(0, leftmostColumn).FormatAsString());
                pane.activePane =(ST_Pane.topRight);
            }
            else if (colSplit == 0)
            {
                pane.topLeftCell = (new CellReference(topRow, 0).FormatAsString());
                pane.activePane = (ST_Pane.bottomLeft);
            }
            else
            {
                pane.topLeftCell = (new CellReference(topRow, leftmostColumn).FormatAsString());
                pane.activePane = (ST_Pane.bottomRight);
            }

            ctView.selection = (null);
            CT_Selection sel = ctView.AddNewSelection();
            sel.pane = (pane.activePane);
        }

        /**
         * Creates a new comment for this sheet. You still
         *  need to assign it to a cell though
         *
         * @deprecated since Nov 2009 this method is not compatible with the common SS interfaces,
         * use {@link NPOI.XSSF.usermodel.XSSFDrawing#CreateCellComment
         *  (NPOI.SS.usermodel.ClientAnchor)} instead
         */
        public XSSFComment CreateComment()
        {
            return CreateDrawingPatriarch().createCellComment(new XSSFClientAnchor());
        }

        /**
         * Create a new row within the sheet and return the high level representation
         *
         * @param rownum  row number
         * @return High level {@link XSSFRow} object representing a row in the sheet
         * @see #RemoveRow(NPOI.SS.usermodel.Row)
         */
        public XSSFRow CreateRow(int rownum)
        {
            CT_Row ctRow;
            XSSFRow prev = _rows[rownum];
            if (prev != null)
            {
                ctRow = prev.GetCTRow();
                ctRow.Set(CT_Row.Factory.newInstance());
            }
            else
            {
                if (_rows.IsEmpty() || rownum > _rows.lastKey())
                {
                    // we can append the new row at the end
                    ctRow = worksheet.sheetData.addNewRow();
                }
                else
                {
                    // get number of rows where row index < rownum
                    // --> this tells us where our row should go
                    int idx = _rows.headMap(rownum).Count;
                    ctRow = worksheet.sheetData.insertNewRow(idx);
                }
            }
            XSSFRow r = new XSSFRow(ctRow, this);
            r.SetRowNum(rownum);
            _rows[rownum] = r;
            return r;
        }

        /**
         * Creates a split pane. Any existing freezepane or split pane is overwritten.
         * @param xSplitPos      Horizonatal position of split (in 1/20th of a point).
         * @param ySplitPos      Vertical position of split (in 1/20th of a point).
         * @param topRow        Top row visible in bottom pane
         * @param leftmostColumn   Left column visible in right pane.
         * @param activePane    Active pane.  One of: PANE_LOWER_RIGHT,
         *                      PANE_UPPER_RIGHT, PANE_LOWER_LEFT, PANE_UPPER_LEFT
         * @see NPOI.SS.usermodel.Sheet#PANE_LOWER_LEFT
         * @see NPOI.SS.usermodel.Sheet#PANE_LOWER_RIGHT
         * @see NPOI.SS.usermodel.Sheet#PANE_UPPER_LEFT
         * @see NPOI.SS.usermodel.Sheet#PANE_UPPER_RIGHT
         */
        public void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, int activePane)
        {
            CreateFreezePane(xSplitPos, ySplitPos, leftmostColumn, topRow);
            GetPane().state = (ST_PaneState.split);
            GetPane().activePane = (ST_Pane.Enum.forInt(activePane));
        }

        public XSSFComment GetCellComment(int row, int column)
        {
            if (sheetComments == null)
            {
                return null;
            }

            String ref1 = new CellReference(row, column).FormatAsString();
            CT_Comment ctComment = sheetComments.GetCTComment(ref1);
            if (ctComment == null) return null;

            XSSFVMLDrawing vml = GetVMLDrawing(false);
            return new XSSFComment(sheetComments, ctComment,
                    vml == null ? null : vml.FindCommentShape(row, column));
        }

        public XSSFHyperlink GetHyperlink(int row, int column)
        {
            String ref1 = new CellReference(row, column).FormatAsString();
            foreach (XSSFHyperlink hyperlink in hyperlinks)
            {
                if (hyperlink.GetCellRef().Equals(ref1))
                {
                    return hyperlink;
                }
            }
            return null;
        }

        /**
         * Vertical page break information used for print layout view, page layout view, drawing print breaks
         * in normal view, and for printing the worksheet.
         *
         * @return column indexes of all the vertical page breaks, never <code>null</code>
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public int[] GetColumnBreaks()
        {
            if (!worksheet.IsSetColBreaks() || worksheet.colBreaks.sizeOfBrkArray() == 0)
            {
                return new int[0];
            }

            List<CT_Break> brkArray = worksheet.colBreaks.brk;

            int[] breaks = new int[brkArray.Count];
            for (int i = 0; i < brkArray.Count; i++)
            {
                CT_Break brk = brkArray[i];
                breaks[i] = (int)brk.id - 1;
            }
            return breaks;
        }

        /**
         * Get the actual column width (in units of 1/256th of a character width )
         *
         * <p>
         * Note, the returned  value is always gerater that {@link #GetDefaultColumnWidth()} because the latter does not include margins.
         * Actual column width measured as the number of characters of the maximum digit width of the
         * numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. There are 4 pixels of margin
         * pAdding (two on each side), plus 1 pixel pAdding for the gridlines.
         * </p>
         *
         * @param columnIndex - the column to set (0-based)
         * @return width - the width in units of 1/256th of a character width
         */
        public int GetColumnWidth(int columnIndex)
        {
            CT_Col col = columnHelper.GetColumn(columnIndex, false);
            double width = col == null || !col.IsSetWidth() ? GetDefaultColumnWidth() : col.width;
            return (int)(width * 256);
        }

        /**
         * Get the default column width for the sheet (if the columns do not define their own width) in
         * characters.
         * <p>
         * Note, this value is different from {@link #GetColumnWidth(int)}. The latter is always greater and includes
         * 4 pixels of margin pAdding (two on each side), plus 1 pixel pAdding for the gridlines.
         * </p>
         * @return column width, default value is 8
         */
        public int GetDefaultColumnWidth()
        {
            CT_SheetFormatPr pr = worksheet.sheetFormatPr;
            return pr == null ? 8 : (int)pr.baseColWidth;
        }

        /**
         * Get the default row height for the sheet (if the rows do not define their own height) in
         * twips (1/20 of  a point)
         *
         * @return  default row height
         */
        public short GetDefaultRowHeight()
        {
            return (short)(GetDefaultRowHeightInPoints() * 20);
        }

        /**
         * Get the default row height for the sheet measued in point size (if the rows do not define their own height).
         *
         * @return  default row height in points
         */
        public float GetDefaultRowHeightInPoints()
        {
            CT_SheetFormatPr pr = worksheet.sheetFormatPr;
            return (float)(pr == null ? 0 : pr.defaultRowHeight);
        }

        private CT_SheetFormatPr GetSheetTypeSheetFormatPr()
        {
            return worksheet.IsSetSheetFormatPr() ?
                   worksheet.sheetFormatPr :
                   worksheet.AddNewSheetFormatPr();
        }

        /**
         * Returns the CellStyle that applies to the given
         *  (0 based) column, or null if no style has been
         *  set for that column
         */
        public ICellStyle GetColumnStyle(int column)
        {
            int idx = columnHelper.GetColDefaultStyle(column);
            return GetWorkbook().getCellStyleAt((short)(idx == -1 ? 0 : idx));
        }

        /**
         * Sets whether the worksheet is displayed from right to left instead of from left to right.
         *
         * @param value true for right to left, false otherwise.
         */
        public void SetRightToLeft(bool value)
        {
            CT_SheetView view = GetDefaultSheetView();
            view.rightToLeft = (value);
        }

        /**
         * Whether the text is displayed in right-to-left mode in the window
         *
         * @return whether the text is displayed in right-to-left mode in the window
         */
        public bool IsRightToLeft()
        {
            CT_SheetView view = GetDefaultSheetView();
            return view == null ? false : view.rightToLeft;
        }

        /**
         * Get whether to display the guts or not,
         * default value is true
         *
         * @return bool - guts or no guts
         */
        public bool GetDisplayGuts()
        {
            CT_SheetPr sheetPr = GetSheetTypeSheetPr();
            CT_OutlinePr outlinePr = sheetPr.outlinePr == null ? new CT_OutlinePr() : sheetPr.outlinePr;
            return outlinePr.showOutlineSymbols;
        }

        /**
         * Set whether to display the guts or not
         *
         * @param value - guts or no guts
         */
        public void SetDisplayGuts(bool value)
        {
            CT_SheetPr sheetPr = GetSheetTypeSheetPr();
            CT_OutlinePr outlinePr = sheetPr.outlinePr == null ? sheetPr.AddNewOutlinePr() : sheetPr.outlinePr;
            outlinePr.showOutlineSymbols = (value);
        }

        /**
         * Gets the flag indicating whether the window should show 0 (zero) in cells Containing zero value.
         * When false, cells with zero value appear blank instead of Showing the number zero.
         *
         * @return whether all zero values on the worksheet are displayed
         */
        public bool IsDisplayZeros()
        {
            CT_SheetView view = GetDefaultSheetView();
            return view == null ? true : view.showZeros;
        }

        /**
         * Set whether the window should show 0 (zero) in cells Containing zero value.
         * When false, cells with zero value appear blank instead of Showing the number zero.
         *
         * @param value whether to display or hide all zero values on the worksheet
         */
        public void SetDisplayZeros(bool value)
        {
            CT_SheetView view = GetSheetTypeSheetView();
            view.showZeros = (value);
        }

        /**
         * Gets the first row on the sheet
         *
         * @return the number of the first logical row on the sheet, zero based
         */
        public int GetFirstRowNum()
        {
            return _rows.Count == 0 ? 0 : _rows.firstKey();
        }

        /**
         * Flag indicating whether the Fit to Page print option is enabled.
         *
         * @return <code>true</code>
         */
        public bool GetFitToPage()
        {
            CT_SheetPr sheetPr = GetSheetTypeSheetPr();
            CT_PageSetUpPr psSetup = (sheetPr == null || !sheetPr.IsSetPageSetUpPr()) ?
                    new CT_PageSetUpPr() : sheetPr.pageSetUpPr;
            return psSetup.fitToPage;
        }

        private CT_SheetPr GetSheetTypeSheetPr()
        {
            if (worksheet.sheetPr == null)
            {
                worksheet.sheetPr = new CT_SheetPr();
            }
            return worksheet.sheetPr;
        }

        private CT_HeaderFooter GetSheetTypeHeaderFooter()
        {
            if (worksheet.headerFooter == null)
            {
                worksheet.headerFooter = new CT_HeaderFooter();
            }
            return worksheet.headerFooter;
        }



        /**
         * Returns the default footer for the sheet,
         *  creating one as needed.
         * You may also want to look at
         *  {@link #GetFirstFooter()},
         *  {@link #GetOddFooter()} and
         *  {@link #GetEvenFooter()}
         */
        public IFooter GetFooter()
        {
            // The default footer is an odd footer
            return GetOddFooter();
        }

        /**
         * Returns the default header for the sheet,
         *  creating one as needed.
         * You may also want to look at
         *  {@link #GetFirstHeader()},
         *  {@link #GetOddHeader()} and
         *  {@link #GetEvenHeader()}
         */
        public IHeader GetHeader()
        {
            // The default header is an odd header
            return GetOddHeader();
        }

        /**
         * Returns the odd footer. Used on all pages unless
         *  other footers also present, when used on only
         *  odd pages.
         */
        public IFooter GetOddFooter()
        {
            return new XSSFOddFooter(GetSheetTypeHeaderFooter());
        }
        /**
         * Returns the even footer. Not there by default, but
         *  when Set, used on even pages.
         */
        public IFooter GetEvenFooter()
        {
            return new XSSFEvenFooter(GetSheetTypeHeaderFooter());
        }
        /**
         * Returns the first page footer. Not there by
         *  default, but when Set, used on the first page.
         */
        public IFooter GetFirstFooter()
        {
            return new XSSFFirstFooter(GetSheetTypeHeaderFooter());
        }

        /**
         * Returns the odd header. Used on all pages unless
         *  other headers also present, when used on only
         *  odd pages.
         */
        public IHeader GetOddHeader()
        {
            return new XSSFOddHeader(GetSheetTypeHeaderFooter());
        }
        /**
         * Returns the even header. Not there by default, but
         *  when Set, used on even pages.
         */
        public IHeader GetEvenHeader()
        {
            return new XSSFEvenHeader(GetSheetTypeHeaderFooter());
        }
        /**
         * Returns the first page header. Not there by
         *  default, but when Set, used on the first page.
         */
        public IHeader GetFirstHeader()
        {
            return new XSSFFirstHeader(GetSheetTypeHeaderFooter());
        }


        /**
         * Determine whether printed output for this sheet will be horizontally centered.
         */
        public bool GetHorizontallyCenter()
        {
            CT_PrintOptions opts = worksheet.printOptions;
            return opts != null && opts.horizontalCentered;
        }

        public int GetLastRowNum()
        {
            return _rows.Count == 0 ? 0 : _rows.lastKey();
        }

        public short GetLeftCol()
        {
            String cellRef = worksheet.GetSheetViews().getSheetViewArray(0).getTopLeftCell();
            CellReference cellReference = new CellReference(cellRef);
            return cellReference.GetCol();
        }

        /**
         * Gets the size of the margin in inches.
         *
         * @param margin which margin to get
         * @return the size of the margin
         * @see Sheet#LeftMargin
         * @see Sheet#RightMargin
         * @see Sheet#TopMargin
         * @see Sheet#BottomMargin
         * @see Sheet#HeaderMargin
         * @see Sheet#FooterMargin
         */
        public double GetMargin(short margin)
        {
            if (!worksheet.IsSetPageMargins()) return 0;

            CT_PageMargins pageMargins = worksheet.pageMargins;
            switch (margin)
            {
                case LeftMargin:
                    return pageMargins.left;
                case RightMargin:
                    return pageMargins.right;
                case TopMargin:
                    return pageMargins.top;
                case BottomMargin:
                    return pageMargins.bottom;
                case HeaderMargin:
                    return pageMargins.header;
                case FooterMargin:
                    return pageMargins.footer;
                default:
                    throw new ArgumentException("Unknown margin constant:  " + margin);
            }
        }

        /**
         * Sets the size of the margin in inches.
         *
         * @param margin which margin to get
         * @param size the size of the margin
         * @see Sheet#LeftMargin
         * @see Sheet#RightMargin
         * @see Sheet#TopMargin
         * @see Sheet#BottomMargin
         * @see Sheet#HeaderMargin
         * @see Sheet#FooterMargin
         */
        public void SetMargin(short margin, double size)
        {
            CT_PageMargins pageMargins = worksheet.IsSetPageMargins() ?
                    worksheet.GetPageMargins() : worksheet.AddNewPageMargins();
            switch (margin)
            {
                case LeftMargin:
                    pageMargins.left = (size);
                    break;
                case RightMargin:
                    pageMargins.right = (size);
                    break;
                case TopMargin:
                    pageMargins.top = (size);
                    break;
                case BottomMargin:
                    pageMargins.bottom = (size);
                    break;
                case HeaderMargin:
                    pageMargins.header = (size);
                    break;
                case FooterMargin:
                    pageMargins.footer = (size);
                    break;
                default:
                    throw new ArgumentException("Unknown margin constant:  " + margin);
            }
        }

        /**
         * @return the merged region at the specified index
         * @throws InvalidOperationException if this worksheet does not contain merged regions
         */
        public CellRangeAddress GetMergedRegion(int index)
        {
            CT_MergeCells ctMergeCells = worksheet.GetMergeCells();
            if (ctMergeCells == null) throw new InvalidOperationException("This worksheet does not contain merged regions");

            CT_MergeCell ctMergeCell = ctMergeCells.GetMergeCellArray(index);
            String ref1 = ctMergeCell.GetRef();
            return CellRangeAddress.ValueOf(ref1);
        }

        /**
         * Returns the number of merged regions defined in this worksheet
         *
         * @return number of merged regions in this worksheet
         */
        public int GetNumMergedRegions()
        {
            CT_MergeCells ctMergeCells = worksheet.GetMergeCells();
            return ctMergeCells == null ? 0 : ctMergeCells.sizeOfMergeCellArray();
        }

        public int GetNumHyperlinks()
        {
            return hyperlinks.Count;
        }

        /**
         * Returns the information regarding the currently configured pane (split or freeze).
         *
         * @return null if no pane configured, or the pane information.
         */
        public PaneInformation GetPaneInformation()
        {
            CT_Pane pane = GetDefaultSheetView().getPane();
            // no pane configured
            if (pane == null) return null;

            CellReference cellRef = pane.IsSetTopLeftCell() ? new CellReference(pane.GetTopLeftCell()) : null;
            return new PaneInformation((short)pane.GetXSplit(), (short)pane.GetYSplit(),
                    (short)(cellRef == null ? 0 : cellRef.GetRow()), (cellRef == null ? 0 : cellRef.GetCol()),
                    (byte)(pane.GetActivePane().intValue() - 1), pane.GetState() == ST_PaneState.FROZEN);
        }

        /**
         * Returns the number of phsyically defined rows (NOT the number of rows in the sheet)
         *
         * @return the number of phsyically defined rows
         */
        public int GetPhysicalNumberOfRows()
        {
            return _rows.Count;
        }

        /**
         * Gets the print Setup object.
         *
         * @return The user model for the print Setup object.
         */
        public XSSFPrintSetup GetPrintSetup()
        {
            return new XSSFPrintSetup(worksheet);
        }

        /**
         * Answer whether protection is enabled or disabled
         *
         * @return true => protection enabled; false => protection disabled
         */
        public bool GetProtect()
        {
            return worksheet.IsSetSheetProtection() && sheetProtectionEnabled();
        }

        /**
         * Enables sheet protection and Sets the password for the sheet.
         * Also Sets some attributes on the {@link CT_SheetProtection} that correspond to
         * the default values used by Excel
         * 
         * @param password to set for protection. Pass <code>null</code> to remove protection
         */
        public void protectSheet(String password)
        {

            if (password != null)
            {
                CT_SheetProtection sheetProtection = worksheet.AddNewSheetProtection();
                sheetProtection.password = StringToExcelPassword(password).ToBytes();
                sheetProtection.sheet = (true);
                sheetProtection.scenarios = (true);
                sheetProtection.objects = (true);
            }
            else
            {
                worksheet.UnsetSheetProtection();
            }
        }

        /**
         * Converts a String to a {@link STUnsignedshortHex} value that Contains the {@link PasswordRecord#hashPassword(String)}
         * value in hexadecimal format
         *  
         * @param password the password string you wish convert to an {@link STUnsignedshortHex}
         * @return {@link STUnsignedshortHex} that Contains Excel hashed password in Hex format
         */
        private ST_UnsignedshortHex StringToExcelPassword(String password)
        {
            ST_UnsignedshortHex hexPassword = new ST_UnsignedshortHex();
            hexPassword.StringValue = HexDump.ShortToHex(PasswordRecord.HashPassword(password)).ToString().Substring(2);
            return hexPassword;
        }

        /**
         * Returns the logical row ( 0-based).  If you ask for a row that is not
         * defined you get a null.  This is to say row 4 represents the fifth row on a sheet.
         *
         * @param rownum  row to get
         * @return <code>XSSFRow</code> representing the rownumber or <code>null</code> if its not defined on the sheet
         */
        public XSSFRow GetRow(int rownum)
        {
            return _rows.Get(rownum);
        }

        /**
         * Horizontal page break information used for print layout view, page layout view, drawing print breaks in normal
         *  view, and for printing the worksheet.
         *
         * @return row indexes of all the horizontal page breaks, never <code>null</code>
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public int[] GetRowBreaks()
        {
            if (!worksheet.IsSetRowBreaks() || worksheet.GetRowBreaks().sizeOfBrkArray() == 0)
            {
                return new int[0];
            }

            CT_Break[] brkArray = worksheet.GetRowBreaks().getBrkArray();
            int[] breaks = new int[brkArray.Length];
            for (int i = 0; i < brkArray.Length; i++)
            {
                CT_Break brk = brkArray[i];
                breaks[i] = (int)brk.id - 1;
            }
            return breaks;
        }

        /**
         * Flag indicating whether summary rows appear below detail in an outline, when Applying an outline.
         *
         * <p>
         * When true a summary row is inserted below the detailed data being summarized and a
         * new outline level is established on that row.
         * </p>
         * <p>
         * When false a summary row is inserted above the detailed data being summarized and a new outline level
         * is established on that row.
         * </p>
         * @return <code>true</code> if row summaries appear below detail in the outline
         */
        public bool GetRowSumsBelow()
        {
            CT_SheetPr sheetPr = worksheet.GetSheetPr();
            CT_OutlinePr outlinePr = (sheetPr != null && sheetPr.IsSetOutlinePr())
                    ? sheetPr.GetOutlinePr() : null;
            return outlinePr == null || outlinePr.GetSummaryBelow();
        }

        /**
         * Flag indicating whether summary rows appear below detail in an outline, when Applying an outline.
         *
         * <p>
         * When true a summary row is inserted below the detailed data being summarized and a
         * new outline level is established on that row.
         * </p>
         * <p>
         * When false a summary row is inserted above the detailed data being summarized and a new outline level
         * is established on that row.
         * </p>
         * @param value <code>true</code> if row summaries appear below detail in the outline
         */
        public void SetRowSumsBelow(bool value)
        {
            ensureOutlinePr().SetSummaryBelow(value);
        }

        /**
         * Flag indicating whether summary columns appear to the right of detail in an outline, when Applying an outline.
         *
         * <p>
         * When true a summary column is inserted to the right of the detailed data being summarized
         * and a new outline level is established on that column.
         * </p>
         * <p>
         * When false a summary column is inserted to the left of the detailed data being
         * summarized and a new outline level is established on that column.
         * </p>
         * @return <code>true</code> if col summaries appear right of the detail in the outline
         */
        public bool GetRowSumsRight()
        {
            CT_SheetPr sheetPr = worksheet.GetSheetPr();
            CT_OutlinePr outlinePr = (sheetPr != null && sheetPr.IsSetOutlinePr())
                    ? sheetPr.GetOutlinePr() : CT_OutlinePr.Factory.newInstance();
            return outlinePr.GetSummaryRight();
        }

        /**
         * Flag indicating whether summary columns appear to the right of detail in an outline, when Applying an outline.
         *
         * <p>
         * When true a summary column is inserted to the right of the detailed data being summarized
         * and a new outline level is established on that column.
         * </p>
         * <p>
         * When false a summary column is inserted to the left of the detailed data being
         * summarized and a new outline level is established on that column.
         * </p>
         * @param value <code>true</code> if col summaries appear right of the detail in the outline
         */
        public void SetRowSumsRight(bool value)
        {
            ensureOutlinePr().SetSummaryRight(value);
        }


        /**
         * Ensure CT_Worksheet.CT_SheetPr.CT_OutlinePr
         */
        private CT_OutlinePr ensureOutlinePr()
        {
            CT_SheetPr sheetPr = worksheet.IsSetSheetPr() ? worksheet.GetSheetPr() : worksheet.AddNewSheetPr();
            return sheetPr.IsSetOutlinePr() ? sheetPr.GetOutlinePr() : sheetPr.AddNewOutlinePr();
        }

        /**
         * A flag indicating whether scenarios are locked when the sheet is protected.
         *
         * @return true => protection enabled; false => protection disabled
         */
        public bool GetScenarioProtect()
        {
            return worksheet.IsSetSheetProtection() && worksheet.GetSheetProtection().getScenarios();
        }

        /**
         * The top row in the visible view when the sheet is
         * first viewed after opening it in a viewer
         *
         * @return integer indicating the rownum (0 based) of the top row
         */
        public short GetTopRow()
        {
            String cellRef = GetSheetTypeSheetView().getTopLeftCell();
            CellReference cellReference = new CellReference(cellRef);
            return (short)cellReference.GetRow();
        }

        /**
         * Determine whether printed output for this sheet will be vertically centered.
         *
         * @return whether printed output for this sheet will be vertically centered.
         */
        public bool GetVerticallyCenter()
        {
            CTPrintOptions opts = worksheet.GetPrintOptions();
            return opts != null && opts.GetVerticalCentered();
        }

        /**
         * Group between (0 based) columns
         */
        public void groupColumn(int fromColumn, int toColumn)
        {
            groupColumn1Based(fromColumn + 1, toColumn + 1);
        }
        private void groupColumn1Based(int fromColumn, int toColumn)
        {
            CTCols ctCols = worksheet.GetColsArray(0);
            CT_Col ctCol = CTCol.Factory.newInstance();
            ctCol.SetMin(fromColumn);
            ctCol.SetMax(toColumn);
            this.columnHelper.AddCleanColIntoCols(ctCols, ctCol);
            for (int index = fromColumn; index <= toColumn; index++)
            {
                CT_Col col = columnHelper.GetColumn1Based(index, false);
                //col must exist
                short outlineLevel = col.GetOutlineLevel();
                col.SetOutlineLevel((short)(outlineLevel + 1));
                index = (int)col.GetMax();
            }
            worksheet.SetColsArray(0, ctCols);
            SetSheetFormatPrOutlineLevelCol();
        }

        /**
         * Tie a range of cell toGether so that they can be collapsed or expanded
         *
         * @param fromRow   start row (0-based)
         * @param toRow     end row (0-based)
         */
        public void groupRow(int fromRow, int toRow)
        {
            for (int i = fromRow; i <= toRow; i++)
            {
                XSSFRow xrow = GetRow(i);
                if (xrow == null)
                {
                    xrow = CreateRow(i);
                }
                CT_Row ctrow = xrow.GetCTRow();
                short outlineLevel = ctrow.GetOutlineLevel();
                ctrow.SetOutlineLevel((short)(outlineLevel + 1));
            }
            SetSheetFormatPrOutlineLevelRow();
        }

        private short GetMaxOutlineLevelRows()
        {
            short outlineLevel = 0;
            foreach (XSSFRow xrow in _rows.values())
            {
                outlineLevel = xrow.GetCTRow().getOutlineLevel() > outlineLevel ? xrow.GetCTRow().getOutlineLevel() : outlineLevel;
            }
            return outlineLevel;
        }


        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        private short GetMaxOutlineLevelCols()
        {
            CTCols ctCols = worksheet.GetColsArray(0);
            short outlineLevel = 0;
            foreach (CT_Col col in ctCols.GetColArray())
            {
                outlineLevel = col.GetOutlineLevel() > outlineLevel ? col.GetOutlineLevel() : outlineLevel;
            }
            return outlineLevel;
        }

        /**
         * Determines if there is a page break at the indicated column
         */
        public bool IsColumnBroken(int column)
        {
            int[] colBreaks = GetColumnBreaks();
            for (int i = 0; i < colBreaks.Length; i++)
            {
                if (colBreaks[i] == column)
                {
                    return true;
                }
            }
            return false;
        }

        /**
         * Get the hidden state for a given column.
         *
         * @param columnIndex - the column to set (0-based)
         * @return hidden - <code>false</code> if the column is visible
         */
        public bool IsColumnHidden(int columnIndex)
        {
            CT_Col col = columnHelper.GetColumn(columnIndex, false);
            return col != null && col.GetHidden();
        }

        /**
         * Gets the flag indicating whether this sheet should display formulas.
         *
         * @return <code>true</code> if this sheet should display formulas.
         */
        public bool IsDisplayFormulas()
        {
            return GetSheetTypeSheetView().getShowFormulas();
        }

        /**
         * Gets the flag indicating whether this sheet displays the lines
         * between rows and columns to make editing and Reading easier.
         *
         * @return <code>true</code> if this sheet displays gridlines.
         * @see #isPrintGridlines() to check if printing of gridlines is turned on or off
         */
        public bool IsDisplayGridlines()
        {
            return GetSheetTypeSheetView().getShowGridLines();
        }

        /**
         * Sets the flag indicating whether this sheet should display the lines
         * between rows and columns to make editing and Reading easier.
         * To turn printing of gridlines use {@link #SetPrintGridlines(bool)}
         *
         *
         * @param show <code>true</code> if this sheet should display gridlines.
         * @see #SetPrintGridlines(bool)
         */
        public void SetDisplayGridlines(bool Show)
        {
            GetSheetTypeSheetView().setShowGridLines(show);
        }

        /**
         * Gets the flag indicating whether this sheet should display row and column headings.
         * <p>
         * Row heading are the row numbers to the side of the sheet
         * </p>
         * <p>
         * Column heading are the letters or numbers that appear above the columns of the sheet
         * </p>
         *
         * @return <code>true</code> if this sheet should display row and column headings.
         */
        public bool IsDisplayRowColHeadings()
        {
            return GetSheetTypeSheetView().getShowRowColHeaders();
        }

        /**
         * Sets the flag indicating whether this sheet should display row and column headings.
         * <p>
         * Row heading are the row numbers to the side of the sheet
         * </p>
         * <p>
         * Column heading are the letters or numbers that appear above the columns of the sheet
         * </p>
         *
         * @param show <code>true</code> if this sheet should display row and column headings.
         */
        public void SetDisplayRowColHeadings(bool Show)
        {
            GetSheetTypeSheetView().setShowRowColHeaders(show);
        }

        /**
         * Returns whether gridlines are printed.
         *
         * @return whether gridlines are printed
         */
        public bool IsPrintGridlines()
        {
            CTPrintOptions opts = worksheet.GetPrintOptions();
            return opts != null && opts.GetGridLines();
        }

        /**
         * Turns on or off the printing of gridlines.
         *
         * @param value bool to turn on or off the printing of gridlines
         */
        public void SetPrintGridlines(bool value)
        {
            CTPrintOptions opts = worksheet.IsSetPrintOptions() ?
                    worksheet.GetPrintOptions() : worksheet.AddNewPrintOptions();
            opts.SetGridLines(value);
        }

        /**
         * Tests if there is a page break at the indicated row
         *
         * @param row index of the row to test
         * @return <code>true</code> if there is a page break at the indicated row
         */
        public bool IsRowBroken(int row)
        {
            int[] rowBreaks = GetRowBreaks();
            for (int i = 0; i < rowBreaks.Length; i++)
            {
                if (rowBreaks[i] == row)
                {
                    return true;
                }
            }
            return false;
        }

        /**
         * Sets a page break at the indicated row
         * Breaks occur above the specified row and left of the specified column inclusive.
         *
         * For example, <code>sheet.SetColumnBreak(2);</code> breaks the sheet into two parts
         * with columns A,B,C in the first and D,E,... in the second. Simuilar, <code>sheet.SetRowBreak(2);</code>
         * breaks the sheet into two parts with first three rows (rownum=1...3) in the first part
         * and rows starting with rownum=4 in the second.
         *
         * @param row the row to break, inclusive
         */
        public void SetRowBreak(int row)
        {
            CTPageBreak pgBreak = worksheet.IsSetRowBreaks() ? worksheet.GetRowBreaks() : worksheet.AddNewRowBreaks();
            if (!isRowBroken(row))
            {
                CT_Break brk = pgBreak.AddNewBrk();
                brk.Id = row + 1; // this is id of the row element which is 1-based: <row r="1" ... >
                brk.SetMan(true);
                brk.SetMax(SpreadsheetVersion.EXCEL2007.getLastColumnIndex()); //end column of the break

                pgBreak.SetCount(pgBreak.sizeOfBrkArray());
                pgBreak.SetManualBreakCount(pgBreak.sizeOfBrkArray());
            }
        }

        /**
         * Removes a page break at the indicated column
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void RemoveColumnBreak(int column)
        {
            if (!worksheet.IsSetColBreaks())
            {
                // no breaks
                return;
            }

            CTPageBreak pgBreak = worksheet.GetColBreaks();
            CT_Break[] brkArray = pgBreak.GetBrkArray();
            for (int i = 0; i < brkArray.Length; i++)
            {
                if (brkArray[i].id == (column + 1))
                {
                    pgBreak.RemoveBrk(i);
                }
            }
        }

        /**
         * Removes a merged region of cells (hence letting them free)
         *
         * @param index of the region to unmerge
         */
        public void RemoveMergedRegion(int index)
        {
            CT_MergeCells ctMergeCells = worksheet.GetMergeCells();

            CT_MergeCell[] mergeCellsArray = new CT_MergeCell[ctMergeCells.sizeOfMergeCellArray() - 1];
            for (int i = 0; i < ctMergeCells.sizeOfMergeCellArray(); i++)
            {
                if (i < index)
                {
                    mergeCellsArray[i] = ctMergeCells.GetMergeCellArray(i);
                }
                else if (i > index)
                {
                    mergeCellsArray[i - 1] = ctMergeCells.GetMergeCellArray(i);
                }
            }
            if (mergeCellsArray.Length > 0)
            {
                ctMergeCells.SetMergeCellArray(mergeCellsArray);
            }
            else
            {
                worksheet.unSetMergeCells();
            }
        }

        /**
         * Remove a row from this sheet.  All cells Contained in the row are Removed as well
         *
         * @param row  the row to Remove.
         */
        public void RemoveRow(IRow row)
        {
            if (row.GetSheet() != this)
            {
                throw new ArgumentException("Specified row does not belong to this sheet");
            }
            // collect cells into a temporary array to avoid ConcurrentModificationException
            List<XSSFCell> cellsToDelete = new List<XSSFCell>();
            foreach (ICell cell in row) cellsToDelete.Add((XSSFCell)cell);

            foreach (XSSFCell cell in cellsToDelete) row.RemoveCell(cell);

            int idx = _rows.headMap(row.GetRowNum()).Count;
            _rows.Remove(row.RowNum);
            worksheet.sheetData.RemoveRow(idx);
        }

        /**
         * Removes the page break at the indicated row
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void RemoveRowBreak(int row)
        {
            if (!worksheet.IsSetRowBreaks())
            {
                return;
            }
            CTPageBreak pgBreak = worksheet.GetRowBreaks();
            CT_Break[] brkArray = pgBreak.GetBrkArray();
            for (int i = 0; i < brkArray.Length; i++)
            {
                if (brkArray[i].id == (row + 1))
                {
                    pgBreak.RemoveBrk(i);
                }
            }
        }

        /**
         * Control if Excel should be asked to recalculate all formulas on this sheet
         * when the workbook is opened.
         *
         *  <p>
         *  Calculating the formula values with {@link NPOI.SS.usermodel.FormulaEvaluator} is the
         *  recommended solution, but this may be used for certain cases where
         *  Evaluation in POI is not possible.
         *  </p>
         *
         *  <p>
         *  It is recommended to force recalcuation of formulas on workbook level using
         *  {@link NPOI.SS.usermodel.Workbook#SetForceFormulaRecalculation(bool)}
         *  to ensure that all cross-worksheet formuals and external dependencies are updated.
         *  </p>
         * @param value true if the application will perform a full recalculation of
         * this worksheet values when the workbook is opened
         *
         * @see NPOI.SS.usermodel.Workbook#SetForceFormulaRecalculation(bool)
         */
        public void SetForceFormulaRecalculation(bool value)
        {
            if (worksheet.IsSetSheetCalcPr())
            {
                // Change the current Setting
                CT_SheetCalcPr calc = worksheet.GetSheetCalcPr();
                calc.SetFullCalcOnLoad(value);
            }
            else if (value)
            {
                // Add the Calc block and set it
                CT_SheetCalcPr calc = worksheet.AddNewSheetCalcPr();
                calc.SetFullCalcOnLoad(value);
            }
            else
            {
                // Not Set, requested not, nothing to do
            }
        }

        /**
         * Whether Excel will be asked to recalculate all formulas when the
         *  workbook is opened.  
         */
        public bool GetForceFormulaRecalculation()
        {
            if (worksheet.IsSetSheetCalcPr())
            {
                CT_SheetCalcPr calc = worksheet.GetSheetCalcPr();
                return calc.GetFullCalcOnLoad();
            }
            return false;
        }

        //    /**
        //     * @return an iterator of the PHYSICAL rows.  Meaning the 3rd element may not
        //     * be the third row if say for instance the second row is undefined.
        //     * Call GetRowNum() on each row if you care which one it is.
        //     */
        //    public Iterator<Row> rowIterator() {
        //    return (Iterator<Row>)(Iterator<? : Row>) _rows.values().iterator();
        //}

        //    /**
        //     * Alias for {@link #rowIterator()} to
        //     *  allow foreach loops
        //     */
        //    public Iterator<Row> iterator()
        //    {
        //        return rowIterator();
        //    }

        /**
         * Flag indicating whether the sheet displays Automatic Page Breaks.
         *
         * @return <code>true</code> if the sheet displays Automatic Page Breaks.
         */
        public bool GetAutobreaks()
        {
            CT_SheetPr sheetPr = GetSheetTypeSheetPr();
            CT_PageSetUpPr psSetup = (sheetPr == null || !sheetPr.IsSetPageSetUpPr()) ?
                    CT_PageSetUpPr.Factory.newInstance() : sheetPr.GetPageSetUpPr();
            return psSetup.GetAutoPageBreaks();
        }

        /**
         * Flag indicating whether the sheet displays Automatic Page Breaks.
         *
         * @param value <code>true</code> if the sheet displays Automatic Page Breaks.
         */
        public void SetAutobreaks(bool value)
        {
            CT_SheetPr sheetPr = GetSheetTypeSheetPr();
            CT_PageSetUpPr psSetup = sheetPr.IsSetPageSetUpPr() ? sheetPr.GetPageSetUpPr() : sheetPr.AddNewPageSetUpPr();
            psSetup.SetAutoPageBreaks(value);
        }

        /**
         * Sets a page break at the indicated column.
         * Breaks occur above the specified row and left of the specified column inclusive.
         *
         * For example, <code>sheet.SetColumnBreak(2);</code> breaks the sheet into two parts
         * with columns A,B,C in the first and D,E,... in the second. Simuilar, <code>sheet.SetRowBreak(2);</code>
         * breaks the sheet into two parts with first three rows (rownum=1...3) in the first part
         * and rows starting with rownum=4 in the second.
         *
         * @param column the column to break, inclusive
         */
        public void SetColumnBreak(int column)
        {
            if (!isColumnBroken(column))
            {
                CTPageBreak pgBreak = worksheet.IsSetColBreaks() ? worksheet.GetColBreaks() : worksheet.AddNewColBreaks();
                CT_Break brk = pgBreak.AddNewBrk();
                brk.Id = column + 1;  // this is id of the row element which is 1-based: <row r="1" ... >
                brk.SetMan(true);
                brk.SetMax(SpreadsheetVersion.EXCEL2007.getLastRowIndex()); //end row of the break

                pgBreak.SetCount(pgBreak.sizeOfBrkArray());
                pgBreak.SetManualBreakCount(pgBreak.sizeOfBrkArray());
            }
        }

        public void SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            if (collapsed)
            {
                collapseColumn(columnNumber);
            }
            else
            {
                expandColumn(columnNumber);
            }
        }

        private void collapseColumn(int columnNumber)
        {
            CTCols cols = worksheet.GetColsArray(0);
            CT_Col col = columnHelper.GetColumn(columnNumber, false);
            int colInfoIx = columnHelper.GetIndexOfColumn(cols, col);
            if (colInfoIx == -1)
            {
                return;
            }
            // Find the start of the group.
            int groupStartColInfoIx = FindStartOfColumnOutlineGroup(colInfoIx);

            CT_Col columnInfo = cols.GetColArray(groupStartColInfoIx);

            // Hide all the columns until the end of the group
            int lastColMax = SetGroupHidden(groupStartColInfoIx, columnInfo
                    .GetOutlineLevel(), true);

            // write collapse field
            SetColumn(lastColMax + 1, null, 0, null, null, Boolean.TRUE);

        }

        private void SetColumn(int tarGetColumnIx, short xfIndex, int style,
                int level, Boolean hidden, Boolean collapsed)
        {
            CTCols cols = worksheet.GetColsArray(0);
            CT_Col ci = null;
            int k = 0;
            for (k = 0; k < cols.sizeOfColArray(); k++)
            {
                CT_Col tci = cols.GetColArray(k);
                if (tci.GetMin() >= tarGetColumnIx
                        && tci.GetMax() <= tarGetColumnIx)
                {
                    ci = tci;
                    break;
                }
                if (tci.GetMin() > tarGetColumnIx)
                {
                    // call column infos after k are for later columns
                    break; // exit now so k will be the correct insert pos
                }
            }

            if (ci == null)
            {
                // okay so there ISN'T a column info record that covers this column
                // so lets create one!
                CT_Col nci = CTCol.Factory.newInstance();
                nci.SetMin(targetColumnIx);
                nci.SetMax(targetColumnIx);
                unSetCollapsed(collapsed, nci);
                this.columnHelper.AddCleanColIntoCols(cols, nci);
                return;
            }

            bool styleChanged = style != null
            && ci.GetStyle() != style;
            bool levelChanged = level != null
            && ci.GetOutlineLevel() != level;
            bool hiddenChanged = hidden != null
            && ci.GetHidden() != hidden;
            bool collapsedChanged = collapsed != null
            && ci.GetCollapsed() != collapsed;
            bool columnChanged = levelChanged || hiddenChanged
            || collapsedChanged || styleChanged;
            if (!columnChanged)
            {
                // do nothing...nothing Changed.
                return;
            }

            if (ci.GetMin() == tarGetColumnIx && ci.GetMax() == tarGetColumnIx)
            {
                // ColumnInfo ci for a single column, the target column
                unSetCollapsed(collapsed, ci);
                return;
            }

            if (ci.GetMin() == tarGetColumnIx || ci.GetMax() == tarGetColumnIx)
            {
                // The target column is at either end of the multi-column ColumnInfo
                // ci
                // we'll just divide the info and create a new one
                if (ci.GetMin() == tarGetColumnIx)
                {
                    ci.SetMin(targetColumnIx + 1);
                }
                else
                {
                    ci.SetMax(targetColumnIx - 1);
                    k++; // adjust insert pos to insert after
                }
                CT_Col nci = columnHelper.CloneCol(cols, ci);
                nci.SetMin(targetColumnIx);
                unSetCollapsed(collapsed, nci);
                this.columnHelper.AddCleanColIntoCols(cols, nci);

            }
            else
            {
                // split to 3 records
                CT_Col ciStart = ci;
                CT_Col ciMid = columnHelper.CloneCol(cols, ci);
                CT_Col ciEnd = columnHelper.CloneCol(cols, ci);
                int lastcolumn = (int)ci.GetMax();

                ciStart.SetMax(targetColumnIx - 1);

                ciMid.SetMin(targetColumnIx);
                ciMid.SetMax(targetColumnIx);
                unSetCollapsed(collapsed, ciMid);
                this.columnHelper.AddCleanColIntoCols(cols, ciMid);

                ciEnd.SetMin(targetColumnIx + 1);
                ciEnd.SetMax(lastcolumn);
                this.columnHelper.AddCleanColIntoCols(cols, ciEnd);
            }
        }

        private void unSetCollapsed(bool collapsed, CT_Col ci)
        {
            if (collapsed)
            {
                ci.SetCollapsed(collapsed);
            }
            else
            {
                ci.unSetCollapsed();
            }
        }

        /**
         * Sets all adjacent columns of the same outline level to the specified
         * hidden status.
         *
         * @param pIdx
         *                the col info index of the start of the outline group
         * @return the column index of the last column in the outline group
         */
        private int SetGroupHidden(int pIdx, int level, bool hidden)
        {
            CTCols cols = worksheet.GetColsArray(0);
            int idx = pIdx;
            CT_Col columnInfo = cols.GetColArray(idx);
            while (idx < cols.sizeOfColArray())
            {
                columnInfo.SetHidden(hidden);
                if (idx + 1 < cols.sizeOfColArray())
                {
                    CT_Col nextColumnInfo = cols.GetColArray(idx + 1);

                    if (!isAdjacentBefore(columnInfo, nextColumnInfo))
                    {
                        break;
                    }

                    if (nextColumnInfo.GetOutlineLevel() < level)
                    {
                        break;
                    }
                    columnInfo = nextColumnInfo;
                }
                idx++;
            }
            return (int)columnInfo.GetMax();
        }

        private bool IsAdjacentBefore(CT_Col col, CT_Col other_col)
        {
            return (col.GetMax() == (other_col.GetMin() - 1));
        }

        private int FindStartOfColumnOutlineGroup(int pIdx)
        {
            // Find the start of the group.
            CTCols cols = worksheet.GetColsArray(0);
            CT_Col columnInfo = cols.GetColArray(pIdx);
            int level = columnInfo.GetOutlineLevel();
            int idx = pIdx;
            while (idx != 0)
            {
                CT_Col prevColumnInfo = cols.GetColArray(idx - 1);
                if (!isAdjacentBefore(prevColumnInfo, columnInfo))
                {
                    break;
                }
                if (prevColumnInfo.GetOutlineLevel() < level)
                {
                    break;
                }
                idx--;
                columnInfo = prevColumnInfo;
            }
            return idx;
        }

        private int FindEndOfColumnOutlineGroup(int colInfoIndex)
        {
            CTCols cols = worksheet.GetColsArray(0);
            // Find the end of the group.
            CT_Col columnInfo = cols.GetColArray(colInfoIndex);
            int level = columnInfo.GetOutlineLevel();
            int idx = colInfoIndex;
            while (idx < cols.sizeOfColArray() - 1)
            {
                CT_Col nextColumnInfo = cols.GetColArray(idx + 1);
                if (!isAdjacentBefore(columnInfo, nextColumnInfo))
                {
                    break;
                }
                if (nextColumnInfo.GetOutlineLevel() < level)
                {
                    break;
                }
                idx++;
                columnInfo = nextColumnInfo;
            }
            return idx;
        }

        private void expandColumn(int columnIndex)
        {
            CTCols cols = worksheet.GetColsArray(0);
            CT_Col col = columnHelper.GetColumn(columnIndex, false);
            int colInfoIx = columnHelper.GetIndexOfColumn(cols, col);

            int idx = FindColInfoIdx((int)col.GetMax(), colInfoIx);
            if (idx == -1)
            {
                return;
            }

            // If it is already expanded do nothing.
            if (!isColumnGroupCollapsed(idx))
            {
                return;
            }

            // Find the start/end of the group.
            int startIdx = FindStartOfColumnOutlineGroup(idx);
            int endIdx = FindEndOfColumnOutlineGroup(idx);

            // expand:
            // colapsed bit must be unset
            // hidden bit Gets unset _if_ surrounding groups are expanded you can
            // determine
            // this by looking at the hidden bit of the enclosing group. You will
            // have
            // to look at the start and the end of the current group to determine
            // which
            // is the enclosing group
            // hidden bit only is altered for this outline level. ie. don't
            // uncollapse Contained groups
            CT_Col columnInfo = cols.GetColArray(endIdx);
            if (!isColumnGroupHiddenByParent(idx))
            {
                int outlineLevel = columnInfo.GetOutlineLevel();
                bool nestedGroup = false;
                for (int i = startIdx; i <= endIdx; i++)
                {
                    CT_Col ci = cols.GetColArray(i);
                    if (outlineLevel == ci.GetOutlineLevel())
                    {
                        ci.unSetHidden();
                        if (nestedGroup)
                        {
                            nestedGroup = false;
                            ci.SetCollapsed(true);
                        }
                    }
                    else
                    {
                        nestedGroup = true;
                    }
                }
            }
            // Write collapse flag (stored in a single col info record after this
            // outline group)
            SetColumn((int)columnInfo.GetMax() + 1, null, null, null,
                    Boolean.FALSE, Boolean.FALSE);
        }

        private bool IsColumnGroupHiddenByParent(int idx)
        {
            CTCols cols = worksheet.GetColsArray(0);
            // Look out outline details of end
            int endLevel = 0;
            bool endHidden = false;
            int endOfOutlineGroupIdx = FindEndOfColumnOutlineGroup(idx);
            if (endOfOutlineGroupIdx < cols.sizeOfColArray())
            {
                CT_Col nextInfo = cols.GetColArray(endOfOutlineGroupIdx + 1);
                if (isAdjacentBefore(cols.GetColArray(endOfOutlineGroupIdx),
                        nextInfo))
                {
                    endLevel = nextInfo.GetOutlineLevel();
                    endHidden = nextInfo.GetHidden();
                }
            }
            // Look out outline details of start
            int startLevel = 0;
            bool startHidden = false;
            int startOfOutlineGroupIdx = FindStartOfColumnOutlineGroup(idx);
            if (startOfOutlineGroupIdx > 0)
            {
                CT_Col prevInfo = cols.GetColArray(startOfOutlineGroupIdx - 1);

                if (isAdjacentBefore(prevInfo, cols
                        .GetColArray(startOfOutlineGroupIdx)))
                {
                    startLevel = prevInfo.GetOutlineLevel();
                    startHidden = prevInfo.GetHidden();
                }

            }
            if (endLevel > startLevel)
            {
                return endHidden;
            }
            return startHidden;
        }

        private int FindColInfoIdx(int columnValue, int fromColInfoIdx)
        {
            CTCols cols = worksheet.GetColsArray(0);

            if (columnValue < 0)
            {
                throw new ArgumentException(
                        "column parameter out of range: " + columnValue);
            }
            if (fromColInfoIdx < 0)
            {
                throw new ArgumentException(
                        "fromIdx parameter out of range: " + fromColInfoIdx);
            }

            for (int k = fromColInfoIdx; k < cols.sizeOfColArray(); k++)
            {
                CT_Col ci = cols.GetColArray(k);

                if (ContainsColumn(ci, columnValue))
                {
                    return k;
                }

                if (ci.GetMin() > fromColInfoIdx)
                {
                    break;
                }

            }
            return -1;
        }

        private bool ContainsColumn(CT_Col col, int columnIndex)
        {
            return col.GetMin() <= columnIndex && columnIndex <= col.GetMax();
        }

        /**
         * 'Collapsed' state is stored in a single column col info record
         * immediately after the outline group
         *
         * @param idx
         * @return a bool represented if the column is collapsed
         */
        private bool IsColumnGroupCollapsed(int idx)
        {
            CTCols cols = worksheet.GetColsArray(0);
            int endOfOutlineGroupIdx = FindEndOfColumnOutlineGroup(idx);
            int nextColInfoIx = endOfOutlineGroupIdx + 1;
            if (nextColInfoIx >= cols.sizeOfColArray())
            {
                return false;
            }
            CT_Col nextColInfo = cols.GetColArray(nextColInfoIx);

            CT_Col col = cols.GetColArray(endOfOutlineGroupIdx);
            if (!isAdjacentBefore(col, nextColInfo))
            {
                return false;
            }

            return nextColInfo.GetCollapsed();
        }

        /**
         * Get the visibility state for a given column.
         *
         * @param columnIndex - the column to get (0-based)
         * @param hidden - the visiblity state of the column
         */
        public void SetColumnHidden(int columnIndex, bool hidden)
        {
            columnHelper.SetColHidden(columnIndex, hidden);
        }

        /**
         * Set the width (in units of 1/256th of a character width)
         *
         * <p>
         * The maximum column width for an individual cell is 255 characters.
         * This value represents the number of characters that can be displayed
         * in a cell that is formatted with the standard font (first font in the workbook).
         * </p>
         *
         * <p>
         * Character width is defined as the maximum digit width
         * of the numbers <code>0, 1, 2, ... 9</code> as rendered
         * using the default font (first font in the workbook).
         * <br/>
         * Unless you are using a very special font, the default character is '0' (zero),
         * this is true for Arial (default font font in HSSF) and Calibri (default font in XSSF)
         * </p>
         *
         * <p>
         * Please note, that the width set by this method includes 4 pixels of margin pAdding (two on each side),
         * plus 1 pixel pAdding for the gridlines (Section 3.3.1.12 of the OOXML spec).
         * This results is a slightly less value of visible characters than passed to this method (approx. 1/2 of a character).
         * </p>
         * <p>
         * To compute the actual number of visible characters,
         *  Excel uses the following formula (Section 3.3.1.12 of the OOXML spec):
         * </p>
         * <code>
         *     width = TRuncate([{Number of Visible Characters} *
         *      {Maximum Digit Width} + {5 pixel pAdding}]/{Maximum Digit Width}*256)/256
         * </code>
         * <p>Using the Calibri font as an example, the maximum digit width of 11 point font size is 7 pixels (at 96 dpi).
         *  If you set a column width to be eight characters wide, e.g. <code>SetColumnWidth(columnIndex, 8*256)</code>,
         *  then the actual value of visible characters (the value Shown in Excel) is derived from the following equation:
         *  <code>
                TRuncate([numChars*7+5]/7*256)/256 = 8;
         *  </code>
         *
         *  which gives <code>7.29</code>.
         *
         * @param columnIndex - the column to set (0-based)
         * @param width - the width in units of 1/256th of a character width
         * @throws ArgumentException if width > 255*256 (the maximum column width in Excel is 255 characters)
         */
        public void SetColumnWidth(int columnIndex, int width)
        {
            if (width > 255 * 256) throw new ArgumentException("The maximum column width for an individual cell is 255 characters.");

            columnHelper.SetColWidth(columnIndex, (double)width / 256);
            columnHelper.SetCustomWidth(columnIndex, true);
        }

        public void SetDefaultColumnStyle(int column, ICellStyle style)
        {
            columnHelper.SetColDefaultStyle(column, style);
        }

        /**
         * Specifies the number of characters of the maximum digit width of the normal style's font.
         * This value does not include margin pAdding or extra pAdding for gridlines. It is only the
         * number of characters.
         *
         * @param width the number of characters. Default value is <code>8</code>.
         */
        public void SetDefaultColumnWidth(int width)
        {
            GetSheetTypeSheetFormatPr().setBaseColWidth(width);
        }

        /**
         * Set the default row height for the sheet (if the rows do not define their own height) in
         * twips (1/20 of  a point)
         *
         * @param  height default row height in  twips (1/20 of  a point)
         */
        public void SetDefaultRowHeight(short height)
        {
            GetSheetTypeSheetFormatPr().setDefaultRowHeight((double)height / 20);

        }

        /**
         * Sets default row height measured in point size.
         *
         * @param height default row height measured in point size.
         */
        public void SetDefaultRowHeightInPoints(float height)
        {
            GetSheetTypeSheetFormatPr().setDefaultRowHeight(height);

        }

        /**
         * Sets the flag indicating whether this sheet should display formulas.
         *
         * @param show <code>true</code> if this sheet should display formulas.
         */
        public void SetDisplayFormulas(bool Show)
        {
            GetSheetTypeSheetView().setShowFormulas(show);
        }

        private CT_SheetView GetSheetTypeSheetView()
        {
            if (GetDefaultSheetView() == null)
            {
                GetSheetTypeSheetViews().setSheetViewArray(0, CT_SheetView.Factory.newInstance());
            }
            return GetDefaultSheetView();
        }

        /**
         * Flag indicating whether the Fit to Page print option is enabled.
         *
         * @param b <code>true</code> if the Fit to Page print option is enabled.
         */
        public void SetFitToPage(bool b)
        {
            GetSheetTypePageSetUpPr().setFitToPage(b);
        }

        /**
         * Center on page horizontally when printing.
         *
         * @param value whether to center on page horizontally when printing.
         */
        public void SetHorizontallyCenter(bool value)
        {
            CTPrintOptions opts = worksheet.IsSetPrintOptions() ?
                    worksheet.GetPrintOptions() : worksheet.AddNewPrintOptions();
            opts.SetHorizontalCentered(value);
        }

        /**
         * Whether the output is vertically centered on the page.
         *
         * @param value true to vertically center, false otherwise.
         */
        public void SetVerticallyCenter(bool value)
        {
            CTPrintOptions opts = worksheet.IsSetPrintOptions() ?
                    worksheet.GetPrintOptions() : worksheet.AddNewPrintOptions();
            opts.SetVerticalCentered(value);
        }

        /**
         * group the row It is possible for collapsed to be false and yet still have
         * the rows in question hidden. This can be achieved by having a lower
         * outline level collapsed, thus hiding all the child rows. Note that in
         * this case, if the lowest level were expanded, the middle level would
         * remain collapsed.
         *
         * @param rowIndex -
         *                the row involved, 0 based
         * @param collapse -
         *                bool value for collapse
         */
        public void SetRowGroupCollapsed(int rowIndex, bool collapse)
        {
            if (collapse)
            {
                collapseRow(rowIndex);
            }
            else
            {
                expandRow(rowIndex);
            }
        }

        /**
         * @param rowIndex the zero based row index to collapse
         */
        private void collapseRow(int rowIndex)
        {
            XSSFRow row = GetRow(rowIndex);
            if (row != null)
            {
                int startRow = FindStartOfRowOutlineGroup(rowIndex);

                // Hide all the columns until the end of the group
                int lastRow = WriteHidden(row, startRow, true);
                if (GetRow(lastRow) != null)
                {
                    GetRow(lastRow).getCT_Row().setCollapsed(true);
                }
                else
                {
                    XSSFRow newRow = CreateRow(lastRow);
                    newRow.GetCTRow().setCollapsed(true);
                }
            }
        }

        /**
         * @param rowIndex the zero based row index to find from
         */
        private int FindStartOfRowOutlineGroup(int rowIndex)
        {
            // Find the start of the group.
            int level = GetRow(rowIndex).getCT_Row().getOutlineLevel();
            int currentRow = rowIndex;
            while (GetRow(currentRow) != null)
            {
                if (GetRow(currentRow).getCT_Row().getOutlineLevel() < level)
                    return currentRow + 1;
                currentRow--;
            }
            return currentRow;
        }

        private int WriteHidden(XSSFRow xRow, int rowIndex, bool hidden)
        {
            int level = xRow.GetCTRow().getOutlineLevel();
            for (Iterator<Row> it = rowIterator(); it.HasNext(); )
            {
                xRow = (XSSFRow)it.next();
                if (xRow.GetCTRow().getOutlineLevel() >= level)
                {
                    xRow.GetCTRow().setHidden(hidden);
                    rowIndex++;
                }

            }
            return rowIndex;
        }

        /**
         * @param rowNumber the zero based row index to expand
         */
        private void expandRow(int rowNumber)
        {
            if (rowNumber == -1)
                return;
            XSSFRow row = GetRow(rowNumber);
            // If it is already expanded do nothing.
            if (!row.GetCTRow().IsSetHidden())
                return;

            // Find the start of the group.
            int startIdx = FindStartOfRowOutlineGroup(rowNumber);

            // Find the end of the group.
            int endIdx = FindEndOfRowOutlineGroup(rowNumber);

            // expand:
            // collapsed must be unset
            // hidden bit Gets unset _if_ surrounding groups are expanded you can
            // determine
            // this by looking at the hidden bit of the enclosing group. You will
            // have
            // to look at the start and the end of the current group to determine
            // which
            // is the enclosing group
            // hidden bit only is altered for this outline level. ie. don't
            // un-collapse Contained groups
            if (!isRowGroupHiddenByParent(rowNumber))
            {
                for (int i = startIdx; i < endIdx; i++)
                {
                    if (row.GetCTRow().getOutlineLevel() == GetRow(i).getCT_Row()
                            .GetOutlineLevel())
                    {
                        GetRow(i).getCT_Row().unsetHidden();
                    }
                    else if (!isRowGroupCollapsed(i))
                    {
                        GetRow(i).getCT_Row().unsetHidden();
                    }
                }
            }
            // Write collapse field
            GetRow(endIdx).getCT_Row().unsetCollapsed();
        }

        /**
         * @param row the zero based row index to find from
         */
        public int FindEndOfRowOutlineGroup(int row)
        {
            int level = GetRow(row).getCT_Row().getOutlineLevel();
            int currentRow;
            for (currentRow = row; currentRow < GetLastRowNum(); currentRow++)
            {
                if (GetRow(currentRow) == null
                        || GetRow(currentRow).getCT_Row().getOutlineLevel() < level)
                {
                    break;
                }
            }
            return currentRow;
        }

        /**
         * @param row the zero based row index to find from
         */
        private bool IsRowGroupHiddenByParent(int row)
        {
            // Look out outline details of end
            int endLevel;
            bool endHidden;
            int endOfOutlineGroupIdx = FindEndOfRowOutlineGroup(row);
            if (GetRow(endOfOutlineGroupIdx) == null)
            {
                endLevel = 0;
                endHidden = false;
            }
            else
            {
                endLevel = GetRow(endOfOutlineGroupIdx).getCT_Row().getOutlineLevel();
                endHidden = GetRow(endOfOutlineGroupIdx).getCT_Row().getHidden();
            }

            // Look out outline details of start
            int startLevel;
            bool startHidden;
            int startOfOutlineGroupIdx = FindStartOfRowOutlineGroup(row);
            if (startOfOutlineGroupIdx < 0
                    || GetRow(startOfOutlineGroupIdx) == null)
            {
                startLevel = 0;
                startHidden = false;
            }
            else
            {
                startLevel = GetRow(startOfOutlineGroupIdx).getCT_Row()
                .GetOutlineLevel();
                startHidden = GetRow(startOfOutlineGroupIdx).getCT_Row()
                .GetHidden();
            }
            if (endLevel > startLevel)
            {
                return endHidden;
            }
            return startHidden;
        }

        /**
         * @param row the zero based row index to find from
         */
        private bool IsRowGroupCollapsed(int row)
        {
            int collapseRow = FindEndOfRowOutlineGroup(row) + 1;
            if (GetRow(collapseRow) == null)
            {
                return false;
            }
            return GetRow(collapseRow).getCT_Row().getCollapsed();
        }

        /**
         * Sets the zoom magnication for the sheet.  The zoom is expressed as a
         * fraction.  For example to express a zoom of 75% use 3 for the numerator
         * and 4 for the denominator.
         *
         * @param numerator     The numerator for the zoom magnification.
         * @param denominator   The denominator for the zoom magnification.
         * @see #SetZoom(int)
         */
        public void SetZoom(int numerator, int denominator)
        {
            int zoom = 100 * numerator / denominator;
            SetZoom(zoom);
        }

        /**
         * Window zoom magnification for current view representing percent values.
         * Valid values range from 10 to 400. Horizontal & Vertical scale toGether.
         *
         * For example:
         * <pre>
         * 10 - 10%
         * 20 - 20%
         * ...
         * 100 - 100%
         * ...
         * 400 - 400%
         * </pre>
         *
         * Current view can be Normal, Page Layout, or Page Break Preview.
         *
         * @param scale window zoom magnification
         * @throws ArgumentException if scale is invalid
         */
        public void SetZoom(int scale)
        {
            if (scale < 10 || scale > 400) throw new ArgumentException("Valid scale values range from 10 to 400");
            GetSheetTypeSheetView().setZoomScale(scale);
        }

        /**
         * Shifts rows between startRow and endRow n number of rows.
         * If you use a negative number, it will shift rows up.
         * Code ensures that rows don't wrap around.
         *
         * Calls ShiftRows(startRow, endRow, n, false, false);
         *
         * <p>
         * Additionally Shifts merged regions that are completely defined in these
         * rows (ie. merged 2 cells on a row to be Shifted).
         * @param startRow the row to start Shifting
         * @param endRow the row to end Shifting
         * @param n the number of rows to shift
         */
        public void ShiftRows(int startRow, int endRow, int n)
        {
            ShiftRows(startRow, endRow, n, false, false);
        }

        /**
         * Shifts rows between startRow and endRow n number of rows.
         * If you use a negative number, it will shift rows up.
         * Code ensures that rows don't wrap around
         *
         * <p>
         * Additionally Shifts merged regions that are completely defined in these
         * rows (ie. merged 2 cells on a row to be Shifted).
         * <p>
         * @param startRow the row to start Shifting
         * @param endRow the row to end Shifting
         * @param n the number of rows to shift
         * @param copyRowHeight whether to copy the row height during the shift
         * @param reSetOriginalRowHeight whether to set the original row's height to the default
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool reSetOriginalRowHeight)
        {
            for (Iterator<Row> it = rowIterator(); it.HasNext(); )
            {
                XSSFRow row = (XSSFRow)it.next();
                int rownum = row.GetRowNum();
                if (rownum < startRow) continue;

                if (!copyRowHeight)
                {
                    row.SetHeight((short)-1);
                }

                if (RemoveRow(startRow, endRow, n, rownum))
                {
                    // remove row from worksheet.GetSheetData row array
                    int idx = _rows.headMap(row.GetRowNum()).Count;
                    worksheet.GetSheetData().removeRow(idx);
                    // remove row from _rows
                    it.Remove();
                }
                else if (rownum >= startRow && rownum <= endRow)
                {
                    row.Shift(n);
                }

                if (sheetComments != null)
                {
                    //TODO shift Note's anchor in the associated /xl/drawing/vmlDrawings#.vml
                    CT_CommentList lst = sheetComments.GetCT_Comments().getCommentList();
                    foreach (CT_Comment comment in lst.GetCommentArray())
                    {
                        CellReference ref1 = new CellReference(comment.GetRef());
                        if (ref1.GetRow() == rownum)
                        {
                            ref1 = new CellReference(rownum + n, ref1.GetCol());
                            comment.SetRef(ref1.FormatAsString());
                        }
                    }
                }
            }
            XSSFRowShifter rowShifter = new XSSFRowShifter(this);

            int sheetIndex = GetWorkbook().getSheetIndex(this);
            FormulaShifter Shifter = FormulaShifter.CreateForRowShift(sheetIndex, startRow, endRow, n);

            rowShifter.updateNamedRanges(Shifter);
            rowShifter.updateFormulas(Shifter);
            rowShifter.ShiftMerged(startRow, endRow, n);
            rowShifter.updateConditionalFormatting(Shifter);

            //rebuild the _rows map
            Dictionary<int, XSSFRow> map = new Dictionary<int, XSSFRow>();
            foreach (XSSFRow r in _rows.Values)
            {
                map.Put(r.GetRowNum(), r);
            }
            _rows = map;
        }

        /**
         * Location of the top left visible cell Location of the top left visible cell in the bottom right
         * pane (when in Left-to-Right mode).
         *
         * @param toprow the top row to show in desktop window pane
         * @param leftcol the left column to show in desktop window pane
         */
        public void ShowInPane(short toprow, short leftcol)
        {
            CellReference cellReference = new CellReference(toprow, leftcol);
            String cellRef = cellReference.FormatAsString();
            GetPane().setTopLeftCell(cellRef);
        }

        public void UngroupColumn(int fromColumn, int toColumn)
        {
            CTCols cols = worksheet.GetColsArray(0);
            for (int index = fromColumn; index <= toColumn; index++)
            {
                CT_Col col = columnHelper.GetColumn(index, false);
                if (col != null)
                {
                    short outlineLevel = col.GetOutlineLevel();
                    col.SetOutlineLevel((short)(outlineLevel - 1));
                    index = (int)col.GetMax();

                    if (col.GetOutlineLevel() <= 0)
                    {
                        int colIndex = columnHelper.GetIndexOfColumn(cols, col);
                        worksheet.GetColsArray(0).removeCol(colIndex);
                    }
                }
            }
            worksheet.SetColsArray(0, cols);
            SetSheetFormatPrOutlineLevelCol();
        }

        /**
         * Ungroup a range of rows that were previously groupped
         *
         * @param fromRow   start row (0-based)
         * @param toRow     end row (0-based)
         */
        public void UngroupRow(int fromRow, int toRow)
        {
            for (int i = fromRow; i <= toRow; i++)
            {
                XSSFRow xrow = GetRow(i);
                if (xrow != null)
                {
                    CT_Row ctrow = xrow.GetCTRow();
                    short outlinelevel = ctrow.GetOutlineLevel();
                    ctrow.SetOutlineLevel((short)(outlinelevel - 1));
                    //remove a row only if the row has no cell and if the outline level is 0
                    if (ctrow.GetOutlineLevel() == 0 && xrow.GetFirstCellNum() == -1)
                    {
                        RemoveRow(xrow);
                    }
                }
            }
            SetSheetFormatPrOutlineLevelRow();
        }

        private void SetSheetFormatPrOutlineLevelRow()
        {
            short maxLevelRow = GetMaxOutlineLevelRows();
            GetSheetTypeSheetFormatPr().setOutlineLevelRow(maxLevelRow);
        }

        private void SetSheetFormatPrOutlineLevelCol()
        {
            short maxLevelCol = GetMaxOutlineLevelCols();
            GetSheetTypeSheetFormatPr().setOutlineLevelCol(maxLevelCol);
        }

        private CT_SheetViews GetSheetTypeSheetViews()
        {
            if (worksheet.GetSheetViews() == null)
            {
                worksheet.SetSheetViews(CT_SheetViews.Factory.newInstance());
                worksheet.GetSheetViews().addNewSheetView();
            }
            return worksheet.GetSheetViews();
        }

        /**
         * Returns a flag indicating whether this sheet is selected.
         * <p>
         * When only 1 sheet is selected and active, this value should be in synch with the activeTab value.
         * In case of a conflict, the Start Part Setting wins and Sets the active sheet tab.
         * </p>
         * Note: multiple sheets can be selected, but only one sheet can be active at one time.
         *
         * @return <code>true</code> if this sheet is selected
         */
        public bool IsSelected()
        {
            CT_SheetView view = GetDefaultSheetView();
            return view != null && view.GetTabSelected();
        }

        /**
         * Sets a flag indicating whether this sheet is selected.
         *
         * <p>
         * When only 1 sheet is selected and active, this value should be in synch with the activeTab value.
         * In case of a conflict, the Start Part Setting wins and Sets the active sheet tab.
         * </p>
         * Note: multiple sheets can be selected, but only one sheet can be active at one time.
         *
         * @param value <code>true</code> if this sheet is selected
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void SetSelected(bool value)
        {
            CT_SheetViews views = GetSheetTypeSheetViews();
            foreach (CT_SheetView view in views.GetSheetViewArray())
            {
                view.SetTabSelected(value);
            }
        }

        /**
         * Assign a cell comment to a cell region in this worksheet
         *
         * @param cellRef cell region
         * @param comment the comment to assign
         * @deprecated since Nov 2009 use {@link XSSFCell#SetCellComment(NPOI.SS.usermodel.Comment)} instead
         */

        public static void SetCellComment(String cellRef, XSSFComment comment)
        {
            CellReference cellReference = new CellReference(cellRef);

            comment.SetRow(cellReference.getRow());
            comment.SetColumn(cellReference.getCol());
        }

        /**
         * Register a hyperlink in the collection of hyperlinks on this sheet
         *
         * @param hyperlink the link to add
         */

        public void AddHyperlink(XSSFHyperlink hyperlink)
        {
            hyperlinks.Add(hyperlink);
        }

        /**
         * Return location of the active cell, e.g. <code>A1</code>.
         *
         * @return the location of the active cell.
         */
        public String GetActiveCell()
        {
            return GetSheetTypeSelection().getActiveCell();
        }

        /**
         * Sets location of the active cell
         *
         * @param cellRef the location of the active cell, e.g. <code>A1</code>..
         */
        public void SetActiveCell(String cellRef)
        {
            CT_Selection ctsel = GetSheetTypeSelection();
            ctsel.SetActiveCell(cellRef);
            ctsel.SetSqref(Arrays.asList(cellRef));
        }

        /**
         * Does this sheet have any comments on it? We need to know,
         *  so we can decide about writing it to disk or not
         */
        public bool HasComments()
        {
            if (sheetComments == null) { return false; }
            return (sheetComments.GetNumberOfComments() > 0);
        }

        protected int GetNumberOfComments()
        {
            if (sheetComments == null) { return 0; }
            return sheetComments.GetNumberOfComments();
        }

        private CT_Selection GetSheetTypeSelection()
        {
            if (GetSheetTypeSheetView().sizeOfSelectionArray() == 0)
            {
                GetSheetTypeSheetView().insertNewSelection(0);
            }
            return GetSheetTypeSheetView().getSelectionArray(0);
        }

        /**
         * Return the default sheet view. This is the last one if the sheet's views, according to sec. 3.3.1.83
         * of the OOXML spec: "A single sheet view defInition. When more than 1 sheet view is defined in the file,
         * it means that when opening the workbook, each sheet view corresponds to a separate window within the
         * spreadsheet application, where each window is Showing the particular sheet. Containing the same
         * workbookViewId value, the last sheetView defInition is loaded, and the others are discarded.
         * When multiple windows are viewing the same sheet, multiple sheetView elements (with corresponding
         * workbookView entries) are saved."
         */
        private CT_SheetView GetDefaultSheetView()
        {
            CT_SheetViews views = GetSheetTypeSheetViews();
            int sz = views == null ? 0 : views.sizeOfSheetViewArray();
            if (sz == 0)
            {
                return null;
            }
            return views.GetSheetViewArray(sz - 1);
        }

        /**
         * Returns the sheet's comments object if there is one,
         *  or null if not
         *
         * @param create create a new comments table if it does not exist
         */
        protected CommentsTable GetCommentsTable(bool Create)
        {
            if (sheetComments == null && Create)
            {
                sheetComments = (CommentsTable)CreateRelationship(XSSFRelation.SHEET_COMMENTS, XSSFFactory.GetInstance(), (int)sheet.GetSheetId());
            }
            return sheetComments;
        }

        private CT_PageSetUpPr GetSheetTypePageSetUpPr()
        {
            CT_SheetPr sheetPr = GetSheetTypeSheetPr();
            return sheetPr.IsSetPageSetUpPr() ? sheetPr.GetPageSetUpPr() : sheetPr.AddNewPageSetUpPr();
        }

        private bool RemoveRow(int startRow, int endRow, int n, int rownum)
        {
            if (rownum >= (startRow + n) && rownum <= (endRow + n))
            {
                if (n > 0 && rownum > endRow)
                {
                    return true;
                }
                else if (n < 0 && rownum < startRow)
                {
                    return true;
                }
            }
            return false;
        }

        private CT_Pane GetPane()
        {
            if (GetDefaultSheetView().getPane() == null)
            {
                GetDefaultSheetView().addNewPane();
            }
            return GetDefaultSheetView().getPane();
        }

        /**
         * Return a master shared formula by index
         *
         * @param sid shared group index
         * @return a CT_CellFormula bean holding shared formula or <code>null</code> if not found
         */
        CT_CellFormula GetSharedFormula(int sid)
        {
            return sharedFormulas.Get(sid);
        }

        void onReadCell(XSSFCell cell)
        {
            //collect cells holding shared formulas
            CTCell ct = cell.GetCTCell();
            CT_CellFormula f = ct.GetF();
            if (f != null && f.GetT() == STCellFormulaType.SHARED && f.IsSetRef() && f.GetStringValue() != null)
            {
                // save a detached  copy to avoid XmlValueDisconnectedException,
                // this may happen when the master cell of a shared formula is Changed
                sharedFormulas.Put((int)f.GetSi(), (CT_CellFormula)f.copy());
            }
            if (f != null && f.GetT() == STCellFormulaType.ARRAY && f.GetRef() != null)
            {
                arrayFormulas.Add(CellRangeAddress.valueOf(f.getRef()));
            }
        }


        protected void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            Write(out1);
            out1.Close();
        }

        protected void Write(Stream out1)
        {

            if (worksheet.sizeOfColsArray() == 1)
            {
                CTCols col = worksheet.GetColsArray(0);
                if (col.sizeOfColArray() == 0)
                {
                    worksheet.SetColsArray(null);
                }
            }

            // Now re-generate our CT_Hyperlinks, if needed
            if (hyperlinks.Count > 0)
            {
                if (worksheet.GetHyperlinks() == null)
                {
                    worksheet.AddNewHyperlinks();
                }
                CT_Hyperlink[] ctHls = new CT_Hyperlink[hyperlinks.Count];
                for (int i = 0; i < ctHls.Length; i++)
                {
                    // If our sheet has hyperlinks, have them add
                    //  any relationships that they might need
                    XSSFHyperlink hyperlink = hyperlinks.Get(i);
                    hyperlink.generateRelationIfNeeded(GetPackagePart());
                    // Now grab their underling object
                    ctHls[i] = hyperlink.GetCT_Hyperlink();
                }
                worksheet.GetHyperlinks().setHyperlinkArray(ctHls);
            }

            foreach (XSSFRow row in _rows.Values)
            {
                row.onDocumentWrite();
            }

            XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            xmlOptions.SetSaveSyntheticDocumentElement(new QName(CT_Worksheet.type.GetName().getNamespaceURI(), "worksheet"));
            Dictionary<String, String> map = new Dictionary<String, String>();
            map.Put(STRelationshipId.type.GetName().getNamespaceURI(), "r");
            xmlOptions.SetSaveSuggestedPrefixes(map);

            worksheet.Save(out1);
        }

        /**
         * @return true when Autofilters are locked and the sheet is protected.
         */
        public bool IsAutoFilterLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getAutoFilter();
        }

        /**
         * @return true when Deleting columns is locked and the sheet is protected.
         */
        public bool IsDeleteColumnsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getDeleteColumns();
        }

        /**
         * @return true when Deleting rows is locked and the sheet is protected.
         */
        public bool IsDeleteRowsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getDeleteRows();
        }

        /**
         * @return true when Formatting cells is locked and the sheet is protected.
         */
        public bool IsFormatCellsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getFormatCells();
        }

        /**
         * @return true when Formatting columns is locked and the sheet is protected.
         */
        public bool IsFormatColumnsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getFormatColumns();
        }

        /**
         * @return true when Formatting rows is locked and the sheet is protected.
         */
        public bool IsFormatRowsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getFormatRows();
        }

        /**
         * @return true when Inserting columns is locked and the sheet is protected.
         */
        public bool IsInsertColumnsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getInsertColumns();
        }

        /**
         * @return true when Inserting hyperlinks is locked and the sheet is protected.
         */
        public bool IsInsertHyperlinksLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getInsertHyperlinks();
        }

        /**
         * @return true when Inserting rows is locked and the sheet is protected.
         */
        public bool IsInsertRowsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getInsertRows();
        }

        /**
         * @return true when Pivot tables are locked and the sheet is protected.
         */
        public bool IsPivotTablesLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getPivotTables();
        }

        /**
         * @return true when Sorting is locked and the sheet is protected.
         */
        public bool IsSortLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getSort();
        }

        /**
         * @return true when Objects are locked and the sheet is protected.
         */
        public bool IsObjectsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getObjects();
        }

        /**
         * @return true when Scenarios are locked and the sheet is protected.
         */
        public bool IsScenariosLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getScenarios();
        }

        /**
         * @return true when Selection of locked cells is locked and the sheet is protected.
         */
        public bool IsSelectLockedCellsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getSelectLockedCells();
        }

        /**
         * @return true when Selection of unlocked cells is locked and the sheet is protected.
         */
        public bool IsSelectUnlockedCellsLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getSelectUnlockedCells();
        }

        /**
         * @return true when Sheet is Protected.
         */
        public bool IsSheetLocked()
        {
            CreateProtectionFieldIfNotPresent();
            return sheetProtectionEnabled() && worksheet.GetSheetProtection().getSheet();
        }

        /**
         * Enable sheet protection
         */
        public void enableLocking()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setSheet(true);
        }

        /**
         * Disable sheet protection
         */
        public void disableLocking()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setSheet(false);
        }

        /**
         * Enable Autofilters locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockAutoFilter()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setAutoFilter(true);
        }

        /**
         * Enable Deleting columns locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockDeleteColumns()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setDeleteColumns(true);
        }

        /**
         * Enable Deleting rows locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockDeleteRows()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setDeleteRows(true);
        }

        /**
         * Enable Formatting cells locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockFormatCells()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setDeleteColumns(true);
        }

        /**
         * Enable Formatting columns locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockFormatColumns()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setFormatColumns(true);
        }

        /**
         * Enable Formatting rows locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockFormatRows()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setFormatRows(true);
        }

        /**
         * Enable Inserting columns locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockInsertColumns()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setInsertColumns(true);
        }

        /**
         * Enable Inserting hyperlinks locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockInsertHyperlinks()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setInsertHyperlinks(true);
        }

        /**
         * Enable Inserting rows locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockInsertRows()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setInsertRows(true);
        }

        /**
         * Enable Pivot Tables locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockPivotTables()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setPivotTables(true);
        }

        /**
         * Enable Sort locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockSort()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setSort(true);
        }

        /**
         * Enable Objects locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockObjects()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setObjects(true);
        }

        /**
         * Enable Scenarios locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockScenarios()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setScenarios(true);
        }

        /**
         * Enable Selection of locked cells locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockSelectLockedCells()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setSelectLockedCells(true);
        }

        /**
         * Enable Selection of unlocked cells locking.
         * This does not modify sheet protection status.
         * To enforce this locking, call {@link #enableLocking()}
         */
        public void lockSelectUnlockedCells()
        {
            CreateProtectionFieldIfNotPresent();
            worksheet.GetSheetProtection().setSelectUnlockedCells(true);
        }

        private void CreateProtectionFieldIfNotPresent()
        {
            if (worksheet.GetSheetProtection() == null)
            {
                worksheet.SetSheetProtection(CT_SheetProtection.Factory.newInstance());
            }
        }

        private bool sheetProtectionEnabled()
        {
            return worksheet.GetSheetProtection().getSheet();
        }

        /* namespace */
        bool IsCellInArrayFormulaContext(XSSFCell cell)
        {
            foreach (CellRangeAddress range in arrayFormulas)
            {
                if (range.IsInRange(cell.GetRowIndex(), cell.GetColumnIndex()))
                {
                    return true;
                }
            }
            return false;
        }

        /* namespace */
        XSSFCell GetFirstCellInArrayFormula(XSSFCell cell)
        {
            foreach (CellRangeAddress range in arrayFormulas)
            {
                if (range.IsInRange(cell.GetRowIndex(), cell.GetColumnIndex()))
                {
                    return GetRow(range.getFirstRow()).getCell(range.getFirstColumn());
                }
            }
            return null;
        }

        /**
         * Also Creates cells if they don't exist
         */
        private ICellRange<XSSFCell> GetCellRange(CellRangeAddress range)
        {
            int firstRow = range.FirstRow;
            int firstColumn = range.FirstColumn;
            int lastRow = range.LastRow;
            int lastColumn = range.LastColumn;
            int height = lastRow - firstRow + 1;
            int width = lastColumn - firstColumn + 1;
            List<XSSFCell> temp = new List<XSSFCell>(height * width);
            for (int rowIn = firstRow; rowIn <= lastRow; rowIn++)
            {
                for (int colIn = firstColumn; colIn <= lastColumn; colIn++)
                {
                    XSSFRow row = GetRow(rowIn);
                    if (row == null)
                    {
                        row = CreateRow(rowIn);
                    }
                    XSSFCell cell = row.GetCell(colIn);
                    if (cell == null)
                    {
                        cell = row.CreateCell(colIn);
                    }
                    temp.Add(cell);
                }
            }
            return SSCellRange.Create(firstRow, firstColumn, height, width, temp, typeof(XSSFCell));
        }

        public ICellRange<XSSFCell> SetArrayFormula(String formula, CellRangeAddress range)
        {

            CellRange<XSSFCell> cr = GetCellRange(range);

            XSSFCell mainArrayFormulaCell = cr.GetTopLeftCell();
            mainArrayFormulaCell.SetCellArrayFormula(formula, range);
            arrayFormulas.Add(range);
            return cr;
        }

        public ICellRange<XSSFCell> RemoveArrayFormula(ICell cell)
        {
            if (cell.GetSheet() != this)
            {
                throw new ArgumentException("Specified cell does not belong to this sheet.");
            }
            foreach (CellRangeAddress range in arrayFormulas)
            {
                if (range.IsInRange(cell.GetRowIndex(), cell.GetColumnIndex()))
                {
                    arrayFormulas.Remove(range);
                    CellRange<XSSFCell> cr = GetCellRange(range);
                    foreach (XSSFCell c in cr)
                    {
                        c.SetCellType(Cell.CELL_TYPE_BLANK);
                    }
                    return cr;
                }
            }
            String ref1 = ((XSSFCell)cell).GetCTCell().getR();
            throw new ArgumentException("Cell " + ref1 + " is not part of an array formula.");
        }


        public IDataValidationHelper GetDataValidationHelper()
        {
            return dataValidationHelper;
        }

        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public List<XSSFDataValidation> GetDataValidations()
        {
            List<XSSFDataValidation> xssfValidations = new List<XSSFDataValidation>();
            CTDataValidations dataValidations = this.worksheet.GetDataValidations();
            if (dataValidations != null && dataValidations.GetCount() > 0)
            {
                foreach (CTDataValidation ctDataValidation in dataValidations.GetDataValidationArray())
                {
                    CellRangeAddressList AddressList = new CellRangeAddressList();


                    List<String> sqref = ctDataValidation.GetSqref();
                    foreach (String stRef in sqref)
                    {
                        String[] regions = stRef.split(" ");
                        for (int i = 0; i < regions.Length; i++)
                        {
                            String[] parts = regions[i].split(":");
                            CellReference begin = new CellReference(parts[0]);
                            CellReference end = parts.Length > 1 ? new CellReference(parts[1]) : begin;
                            CellRangeAddress cellRangeAddress = new CellRangeAddress(begin.GetRow(), end.GetRow(), begin.GetCol(), end.GetCol());
                            AddressList.addCellRangeAddress(cellRangeAddress);
                        }
                    }
                    XSSFDataValidation xssfDataValidation = new XSSFDataValidation(AddressList, ctDataValidation);
                    xssfValidations.Add(xssfDataValidation);
                }
            }
            return xssfValidations;
        }

        public void AddValidationData(IDataValidation dataValidation)
        {
            XSSFDataValidation xssfDataValidation = (XSSFDataValidation)dataValidation;
            CTDataValidations dataValidations = worksheet.GetDataValidations();
            if (dataValidations == null)
            {
                dataValidations = worksheet.AddNewDataValidations();
            }
            int currentCount = dataValidations.sizeOfDataValidationArray();
            CTDataValidation newval = dataValidations.AddNewDataValidation();
            newval.Set(xssfDataValidation.getCtDdataValidation());
            dataValidations.SetCount(currentCount + 1);

        }

        public XSSFAutoFilter SetAutoFilter(CellRangeAddress range)
        {
            CTAutoFilter af = worksheet.GetAutoFilter();
            if (af == null) af = worksheet.AddNewAutoFilter();

            CellRangeAddress norm = new CellRangeAddress(range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn);
            String ref1 = norm.FormatAsString();
            af.@ref = (ref1);

            XSSFWorkbook wb = GetWorkbook();
            int sheetIndex = GetWorkbook().getSheetIndex(this);
            XSSFName name = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
            if (name == null)
            {
                name = wb.CreateBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
                name.GetCTName().setHidden(true);
                CellReference r1 = new CellReference(GetSheetName(), range.FirstRow, range.FirstColumn, true, true);
                CellReference r2 = new CellReference(null, range.LastRow, range.LastColumn, true, true);
                String fmla = r1.FormatAsString() + ":" + r2.FormatAsString();
                name.SetRefersToFormula(fmla);
            }

            return new XSSFAutoFilter(this);
        }

        /**
         * Creates a new Table, and associates it with this Sheet
         */
        public XSSFTable CreateTable()
        {
            if (!worksheet.IsSetTableParts())
            {
                worksheet.AddNewTableParts();
            }

            CTTableParts tblParts = worksheet.GetTableParts();
            CTTablePart tbl = tblParts.AddNewTablePart();

            // Table numbers need to be unique in the file, not just
            //  unique within the sheet. Find the next one
            int tableNumber = GetPackagePart().getPackage().getPartsByContentType(XSSFRelation.TABLE.getContentType()).Count + 1;
            XSSFTable table = (XSSFTable)CreateRelationship(XSSFRelation.TABLE, XSSFFactory.GetInstance(), tableNumber);
            tbl.Id = table.getPackageRelationship().id;

            tables.Put(tbl.GetId(), table);

            return table;
        }

        /**
         * Returns any tables associated with this Sheet
         */
        public List<XSSFTable> GetTables()
        {
            List<XSSFTable> tableList = new List<XSSFTable>(
                  tables.values()
            );
            return tableList;
        }

        public XSSFSheetConditionalFormatting GetSheetConditionalFormatting()
        {
            return new XSSFSheetConditionalFormatting(this);
        }
    }

}