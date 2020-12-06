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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Xml;
using NPOI.HSSF.Record;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel.Helpers;
using NPOI.POIFS.Crypt;

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

        private static double DEFAULT_ROW_HEIGHT = 15.0;
        private static double DEFAULT_MARGIN_HEADER = 0.3;
        private static double DEFAULT_MARGIN_FOOTER = 0.3;
        private static double DEFAULT_MARGIN_TOP = 0.75;
        private static double DEFAULT_MARGIN_BOTTOM = 0.75;
        private static double DEFAULT_MARGIN_LEFT = 0.7;
        private static double DEFAULT_MARGIN_RIGHT = 0.7;
        public static int TWIPS_PER_POINT = 20;

        //TODO make the two variable below private!
        internal CT_Sheet sheet;
        internal CT_Worksheet worksheet;

        private SortedList<int, XSSFRow> _rows = new SortedList<int, XSSFRow>();
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
        public XSSFSheet()
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
        protected internal XSSFSheet(PackagePart part)
            : base(part)
        {
            dataValidationHelper = new XSSFDataValidationHelper(this);
        }

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        internal XSSFSheet(PackagePart part, PackageRelationship rel)
            : this(part)
        {
        }

        /**
         * Returns the parent XSSFWorkbook
         *
         * @return the parent XSSFWorkbook
         */
        public IWorkbook Workbook
        {
            get
            {
                return (XSSFWorkbook)GetParent();

            }
        }

        /**
         * Initialize worksheet data when Reading in an exisiting file.
         */

        internal override void OnDocumentRead()
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

        internal virtual void Read(Stream is1)
        {
            try
            {
                XmlDocument doc = ConvertStreamToXml(is1);
                worksheet = WorksheetDocument.Parse(doc, NamespaceManager).GetWorksheet();
            }
            catch (XmlException e)
            {
                throw new POIXMLException(e);
            }

            InitRows(worksheet);
            columnHelper = new ColumnHelper(worksheet);

            // Look for bits we're interested in
            foreach (RelationPart rp in RelationParts)
            {
                POIXMLDocumentPart p = rp.DocumentPart;
                if (p is CommentsTable)
                {
                    sheetComments = (CommentsTable)p;
                    //break;
                }
                if (p is XSSFTable)
                {
                    tables[rp.Relationship.Id] = (XSSFTable)p;
                }
                if (p is XSSFPivotTable)
                {
                    GetWorkbook().PivotTables.Add((XSSFPivotTable)p);
                }
            }

            // Process external hyperlinks for the sheet, if there are any
            InitHyperlinks();
        }

        /**
         * Initialize worksheet data when creating a new sheet.
         */

        internal override void OnDocumentCreate()
        {
            worksheet = NewSheet();
            InitRows(worksheet);
            columnHelper = new ColumnHelper(worksheet);
            hyperlinks = new List<XSSFHyperlink>();
        }

        private void InitRows(CT_Worksheet worksheetParam)
        {
            _rows.Clear();
            tables = new Dictionary<String, XSSFTable>();
            sharedFormulas = new Dictionary<int, CT_CellFormula>();
            arrayFormulas = new List<CellRangeAddress>();
            if (0 < worksheetParam.sheetData.SizeOfRowArray())
            {
                foreach (CT_Row row in worksheetParam.sheetData.row)
                {
                    XSSFRow r = new XSSFRow(row, this);
                    if (!_rows.ContainsKey(r.RowNum))
                        _rows.Add(r.RowNum, r);
                }
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
                foreach (CT_Hyperlink hyperlink in worksheet.hyperlinks.hyperlink)
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
        private static CT_Worksheet NewSheet()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_SheetFormatPr ctFormat = worksheet.AddNewSheetFormatPr();
            ctFormat.defaultRowHeight = DEFAULT_ROW_HEIGHT;

            CT_SheetView ctView = worksheet.AddNewSheetViews().AddNewSheetView();
            ctView.workbookViewId = (0);

            worksheet.AddNewDimension().@ref = "A1";

            worksheet.AddNewSheetData();

            CT_PageMargins ctMargins = worksheet.AddNewPageMargins();
            ctMargins.bottom = DEFAULT_MARGIN_BOTTOM;
            ctMargins.footer = DEFAULT_MARGIN_FOOTER;
            ctMargins.header = DEFAULT_MARGIN_HEADER;
            ctMargins.left = DEFAULT_MARGIN_LEFT;
            ctMargins.right = DEFAULT_MARGIN_RIGHT;
            ctMargins.top = DEFAULT_MARGIN_TOP;

            return worksheet;
        }

        /**
         * Provide access to the CT_Worksheet bean holding this sheet's data
         *
         * @return the CT_Worksheet bean holding this sheet's data
         */

        internal CT_Worksheet GetCTWorksheet()
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
        public String SheetName
        {
            get
            {
                return sheet.name;
            }
        }

        /// <summary>
        /// Adds a merged region of cells on a sheet.
        /// </summary>
        /// <param name="region">region to merge</param>
        /// <returns>index of this region</returns>
        /// <exception cref="System.ArgumentException">if region contains fewer than 2 cells</exception>
        /// <exception cref="InvalidOperationException">if region intersects with an existing merged region
        /// or multi-cell array formula on this sheet</exception>
        public int AddMergedRegion(CellRangeAddress region)
        {
            return AddMergedRegion(region, true);
        }

        /// <summary>
        /// Adds a merged region of cells (hence those cells form one).
        /// Skips validation.It is possible to create overlapping merged regions
        /// or create a merged region that intersects a multi-cell array formula
        /// with this formula, which may result in a corrupt workbook.
        /// </summary>
        /// <param name="region">region to merge</param>
        /// <returns>index of this region</returns>
        /// <exception cref="System.ArgumentException">if region contains fewer than 2 cells</exception>
        public int AddMergedRegionUnsafe(CellRangeAddress region)
        {
            return AddMergedRegion(region, false);
        }

        /// <summary>
        /// Adds a merged region of cells (hence those cells form one).
        /// </summary>
        /// <param name="region">region (rowfrom/colfrom-rowto/colto) to merge</param>
        /// <param name="validate">whether to validate merged region</param>
        /// <returns>index of this region</returns>
        /// <exception cref="InvalidOperationException">if region intersects with a multi-cell array formula or
        /// if region intersects with an existing region on this sheet</exception>
        /// <exception cref="System.ArgumentException">if region contains fewer than 2 cells</exception>
        private int AddMergedRegion(CellRangeAddress region, bool validate)
        {
            if (region.NumberOfCells < 2)
            {
                throw new ArgumentException("Merged region " + region.FormatAsString() + " must contain 2 or more cells");
            }
            region.Validate(SpreadsheetVersion.EXCEL2007);
            if (validate)
            {
                // throw InvalidOperationException if the argument CellRangeAddress intersects with
                // a multi-cell array formula defined in this sheet
                ValidateArrayFormulas(region);
                // Throw InvalidOperationException if the argument CellRangeAddress intersects with
                // a merged region already in this sheet 
                ValidateMergedRegions(region);
            }

            CT_MergeCells ctMergeCells = worksheet.IsSetMergeCells() ? worksheet.mergeCells : worksheet.AddNewMergeCells();
            CT_MergeCell ctMergeCell = ctMergeCells.AddNewMergeCell();
            ctMergeCell.@ref = (region.FormatAsString());
            return ctMergeCells.sizeOfMergeCellArray();
        }
        /**
         * Verify that the candidate region does not intersect with an existing multi-cell array formula in this sheet
         *
         * @param region
         * @throws InvalidOperationException if candidate region intersects an existing array formula in this sheet
         */
        private void ValidateArrayFormulas(CellRangeAddress region)
        {
            // FIXME: this may be faster if it looped over array formulas directly rather than looping over each cell in
            // the region and searching if that cell belongs to an array formula
            int firstRow = region.FirstRow;
            int firstColumn = region.FirstColumn;
            int lastRow = region.LastRow;
            int lastColumn = region.LastColumn;
            // for each cell in sheet, if cell belongs to an array formula, check if merged region intersects array formula cells
            for (int rowIn = firstRow; rowIn <= lastRow; rowIn++)
            {
                IRow row = GetRow(rowIn);
                if (row == null) continue;
                for (int colIn = firstColumn; colIn <= lastColumn; colIn++)
                {
                    ICell cell = row.GetCell(colIn);
                    if (cell == null) continue;

                    if (cell.IsPartOfArrayFormulaGroup)
                    {
                        CellRangeAddress arrayRange = cell.ArrayFormulaRange;
                        if (arrayRange.NumberOfCells > 1 && region.Intersects(arrayRange))
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
         * Verify that none of the merged regions intersect a multi-cell array formula in this sheet
         *
         * @param region
         * @throws InvalidOperationException if candidate region intersects an existing array formula in this sheet
         */
        private void CheckForMergedRegionsIntersectingArrayFormulas()
        {
            foreach (CellRangeAddress region in MergedRegions)
            {
                ValidateArrayFormulas(region);
            }
        }

        /**
         * Verify that candidate region does not intersect with an existing merged region in this sheet
         *
         * @param candidateRegion
         * @throws InvalidOperationException if candidate region intersects an existing merged region in this sheet
         */
        private void ValidateMergedRegions(CellRangeAddress candidateRegion)
        {
            foreach (CellRangeAddress existingRegion in MergedRegions)
            {
                if (existingRegion.Intersects(candidateRegion))
                {
                    throw new InvalidOperationException("Cannot add merged region " + candidateRegion.FormatAsString() +
                            " to sheet because it overlaps with an existing merged region (" + existingRegion.FormatAsString() + ").");
                }
            }
        }

        /**
         * Verify that no merged regions intersect another merged region in this sheet.
         *
         * @throws InvalidOperationException if at least one region intersects with another merged region in this sheet
         */
        private void CheckForIntersectingMergedRegions()
        {
            List<CellRangeAddress> regions = MergedRegions;
            int size = regions.Count;
            for (int i = 0; i < size; i++)
            {
                CellRangeAddress region = regions[i];
                foreach (CellRangeAddress other in regions.Skip(i)) //regions.subList(i+1, regions.size()
                {
                    if (region.Intersects(other))
                    {
                        String msg = "The range " + region.FormatAsString() +
                                    " intersects with another merged region " +
                                    other.FormatAsString() + " in this sheet";
                        throw new InvalidOperationException(msg);
                    }
                }
            }
        }

        /**
         * Verify that merged regions do not intersect multi-cell array formulas and
         * no merged regions intersect another merged region in this sheet.
         *
         * @throws InvalidOperationException if region intersects with a multi-cell array formula
         * @throws InvalidOperationException if at least one region intersects with another merged region in this sheet
         */
        public void ValidateMergedRegions()
        {
            CheckForMergedRegionsIntersectingArrayFormulas();
            CheckForIntersectingMergedRegions();
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
        XSSFDrawing drawing = null;
        /**
         * Return the sheet's existing Drawing, or null if there isn't yet one.
         * 
         * Use {@link #CreateDrawingPatriarch()} to Get or create
         *
         * @return a SpreadsheetML Drawing
         */
        public XSSFDrawing GetDrawingPatriarch()
        {
            CT_Drawing ctDrawing = GetCTDrawing();
            if (ctDrawing != null)
            {
                // Search the referenced Drawing in the list of the sheet's relations
                foreach (RelationPart rp in RelationParts)
                {
                    POIXMLDocumentPart p = rp.DocumentPart;
                    if (p is XSSFDrawing)
                    {
                        XSSFDrawing dr = (XSSFDrawing)p;
                        String drId = rp.Relationship.Id;
                        if (drId.Equals(ctDrawing.id))
                        {
                            return dr;
                        }
                        break;
                    }
                }
                logger.Log(POILogger.ERROR, "Can't find Drawing with id=" + ctDrawing.id + " in the list of the sheet's relationships");
            }
            return null;
        }

        /**
         * Create a new SpreadsheetML Drawing. If this sheet already Contains a Drawing - return that.
         *
         * @return a SpreadsheetML Drawing
         */

        public IDrawing CreateDrawingPatriarch()
        {
            CT_Drawing ctDrawing = GetCTDrawing();
            if (ctDrawing != null)
            {
                return GetDrawingPatriarch();
            }

            //drawingNumber = #drawings.Count + 1
            int DrawingNumber = GetPackagePart().Package.GetPartsByContentType(XSSFRelation.DRAWINGS.ContentType).Count + 1;
            RelationPart rp = CreateRelationship(XSSFRelation.DRAWINGS, XSSFFactory.GetInstance(), DrawingNumber, false);
            XSSFDrawing drawing = rp.DocumentPart as XSSFDrawing;
            String relId = rp.Relationship.Id;

            //add CT_Drawing element which indicates that this sheet Contains Drawing components built on the DrawingML platform.
            //The relationship Id references the part Containing the DrawingML defInitions.
            ctDrawing = worksheet.AddNewDrawing();
            ctDrawing.id = (/*setter*/relId);

            // Return the newly Created Drawing
            return drawing;
        }


        /**
         * Get VML drawing for this sheet (aka 'legacy' drawig)
         *
         * @param autoCreate if true, then a new VML drawing part is Created
         *
         * @return the VML drawing of <code>null</code> if the drawing was not found and autoCreate=false
         */
        internal XSSFVMLDrawing GetVMLDrawing(bool autoCreate)
        {
            XSSFVMLDrawing drawing = null;
            NPOI.OpenXmlFormats.Spreadsheet.CT_LegacyDrawing ctDrawing = GetCTLegacyDrawing();
            if (ctDrawing == null)
            {
                if (autoCreate)
                {
                    //drawingNumber = #drawings.Count + 1
                    int drawingNumber = GetPackagePart().Package.GetPartsByContentType(XSSFRelation.VML_DRAWINGS.ContentType).Count + 1;
                    RelationPart rp = CreateRelationship(XSSFRelation.VML_DRAWINGS, XSSFFactory.GetInstance(), drawingNumber, false);
                    drawing = rp.DocumentPart as XSSFVMLDrawing;
                    String relId = rp.Relationship.Id;

                    //add CT_LegacyDrawing element which indicates that this sheet Contains drawing components built on the drawingML platform.
                    //The relationship Id references the part Containing the drawing defInitions.
                    ctDrawing = worksheet.AddNewLegacyDrawing();
                    ctDrawing.id = relId;
                }
            }
            else
            {
                //search the referenced drawing in the list of the sheet's relations
                String id = ctDrawing.id;
                foreach (RelationPart rp in RelationParts)
                {
                    POIXMLDocumentPart p = rp.DocumentPart;
                    if (p is XSSFVMLDrawing)
                    {
                        XSSFVMLDrawing dr = (XSSFVMLDrawing)p;
                        String drId = rp.Relationship.Id;
                        if (drId.Equals(id))
                        {
                            drawing = dr;
                            break;
                        }
                        // do not break here since drawing has not been found yet (see bug 52425)
                    }
                }
                if (drawing == null)
                {
                    logger.Log(POILogger.ERROR, "Can't find VML drawing with id=" + id + " in the list of the sheet's relationships");
                }
            }
            return drawing;
        }

        protected virtual CT_Drawing GetCTDrawing()
        {
            return worksheet.drawing;
        }
        protected virtual CT_LegacyDrawing GetCTLegacyDrawing()
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

                if (ctView.IsSetPane()) ctView.UnsetPane();
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
                pane.xSplit = (colSplit);
            }
            else
            {

                if (pane.IsSetXSplit()) pane.UnsetXSplit();
            }
            if (rowSplit > 0)
            {
                pane.ySplit = (rowSplit);
            }
            else
            {
                if (pane.IsSetYSplit()) pane.UnsetYSplit();
            }

            pane.state = (ST_PaneState.frozen);
            if (rowSplit == 0)
            {
                pane.topLeftCell = (new CellReference(0, leftmostColumn).FormatAsString());
                pane.activePane = (ST_Pane.topRight);
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

        int GetLastKey(IList<int> keys)
        {
            int i = keys.Count;
            return keys[keys.Count - 1];
        }

        int HeadMapCount(SortedList<int, XSSFRow> rows, int rownum)
        {
            int count = 0;
            foreach (int key in rows.Keys)
            {
                if (key < rownum)
                {
                    count++;
                }
                else
                {
                    break;
                }
            }
            return count;
        }

        /**
         * Create a new row within the sheet and return the high level representation
         *
         * @param rownum  row number
         * @return High level {@link XSSFRow} object representing a row in the sheet
         * @see #RemoveRow(NPOI.SS.usermodel.Row)
         */
        public virtual IRow CreateRow(int rownum)
        {
            CT_Row ctRow;
            XSSFRow prev = _rows.ContainsKey(rownum) ? _rows[rownum] : null;
            if (prev != null)
            {
                // the Cells in an existing row are invalidated on-purpose, in order to clean up correctly, we
                // need to call the remove, so things like ArrayFormulas and CalculationChain updates are done 
                // correctly. 
                // We remove the cell this way as the internal cell-list is changed by the remove call and 
                // thus would cause ConcurrentModificationException otherwise
                while (prev.FirstCellNum != -1)
                {
                    prev.RemoveCell(prev.GetCell(prev.FirstCellNum));
                }
                ctRow = prev.GetCTRow();
                ctRow.Set(new CT_Row());
            }
            else
            {
                if (_rows.Count == 0 || rownum > GetLastKey(_rows.Keys))
                {
                    // we can append the new row at the end
                    ctRow = worksheet.sheetData.AddNewRow();
                }
                else
                {
                    // get number of rows where row index < rownum
                    // --> this tells us where our row should go
                    int idx = HeadMapCount(_rows, rownum);
                    ctRow = worksheet.sheetData.InsertNewRow(idx);
                }
            }
            XSSFRow r = new XSSFRow(ctRow, this);
            r.RowNum = rownum;
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
        public void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane)
        {
            CreateFreezePane(xSplitPos, ySplitPos, leftmostColumn, topRow);
            GetPane().state = (ST_PaneState.split);
            GetPane().activePane = (ST_Pane)(activePane);
        }
        /// <summary>
        /// Returns cell comment for the specified row and column
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <returns>cell comment or <code>null</code> if not found</returns>
        [Obsolete("deprecated as of 2015-11-23 (circa POI 3.14beta1). Use {@link #getCellComment(CellAddress)} instead.")]
        public IComment GetCellComment(int row, int column)
        {
            return GetCellComment(new CellAddress(row, column));
        }
        /// <summary>
        /// Returns cell comment for the specified location
        /// </summary>
        /// <param name="address">cell location</param>
        /// <returns>return cell comment or null if not found</returns>
        public IComment GetCellComment(CellAddress address)
        {
            if (sheetComments == null)
            {
                return null;
            }

            int row = address.Row;
            int column = address.Column;

            CellAddress ref1 = new CellAddress(row, column);
            CT_Comment ctComment = sheetComments.GetCTComment(ref1);
            if (ctComment == null) return null;

            XSSFVMLDrawing vml = GetVMLDrawing(false);
            return new XSSFComment(sheetComments, ctComment,
                    vml == null ? null : vml.FindCommentShape(row, column));
        }
        /// <summary>
        /// Returns all cell comments on this sheet.
        /// </summary>
        /// <returns>return A Dictionary of each Comment in the sheet, keyed on the cell address where the comment is located.</returns>
        public Dictionary<CellAddress, IComment> GetCellComments()
        {
            if (sheetComments == null)
            {
                return new Dictionary<CellAddress, IComment>();
            }
            return sheetComments.GetCellComments();
        }

        /// <summary>
        /// Get a Hyperlink in this sheet anchored at row, column
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns>return hyperlink if there is a hyperlink anchored at row, column; otherwise returns null</returns>
        public IHyperlink GetHyperlink(int row, int column)
        {
            return GetHyperlink(new CellAddress(row, column));
        }

        /// <summary>
        /// Get a Hyperlink in this sheet located in a cell specified by {code addr}
        /// </summary>
        /// <param name="addr">The address of the cell containing the hyperlink</param>
        /// <returns>return hyperlink if there is a hyperlink anchored at {@code addr}; otherwise returns {@code null}</returns>
        public IHyperlink GetHyperlink(CellAddress addr)
        {
            String ref1 = addr.FormatAsString();
            foreach (XSSFHyperlink hyperlink in hyperlinks)
            {
                if (hyperlink.CellRef.Equals(ref1))
                {
                    return hyperlink;
                }
            }
            return null;
        }

        /// <summary>
        /// Get a list of Hyperlinks in this sheet
        /// </summary>
        /// <returns></returns>
        public List<IHyperlink> GetHyperlinkList()
        {
            return (hyperlinks.ToList<IHyperlink>());
        }

        /**
         * Vertical page break information used for print layout view, page layout view, drawing print breaks
         * in normal view, and for printing the worksheet.
         *
         * @return column indexes of all the vertical page breaks, never <code>null</code>
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public int[] ColumnBreaks
        {
            get
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
            double width = (col == null || !col.IsSetWidth()) ? this.DefaultColumnWidth : col.width;
            return (int)(width * 256);
        }

        /**
         * Get the actual column width in pixels
         * 
         * <p>
         * Please note, that this method works correctly only for workbooks
         * with the default font size (Calibri 11pt for .xlsx).
         * </p>
         */
        public float GetColumnWidthInPixels(int columnIndex)
        {
            float widthIn256 = GetColumnWidth(columnIndex);
            return (float)(widthIn256 / 256.0 * XSSFWorkbook.DEFAULT_CHARACTER_WIDTH);
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
        public int DefaultColumnWidth
        {
            get
            {
                CT_SheetFormatPr pr = worksheet.sheetFormatPr;
                return pr == null ? 8 : (int)pr.baseColWidth;
            }
            set
            {
                GetSheetTypeSheetFormatPr().baseColWidth = (uint)value;
            }
        }

        /**
         * Get the default row height for the sheet (if the rows do not define their own height) in
         * twips (1/20 of  a point)
         *
         * @return  default row height
         */
        public short DefaultRowHeight
        {
            get
            {
                return (short)((decimal)DefaultRowHeightInPoints * TWIPS_PER_POINT);
            }
            set
            {
                DefaultRowHeightInPoints = (float)value / TWIPS_PER_POINT;
            }
        }

        /**
         * Get the default row height for the sheet measued in point size (if the rows do not define their own height).
         *
         * @return  default row height in points
         */
        public float DefaultRowHeightInPoints
        {
            get
            {
                CT_SheetFormatPr pr = worksheet.sheetFormatPr;
                return (float)(pr == null ? 0 : pr.defaultRowHeight);
            }
            set
            {
                CT_SheetFormatPr pr = GetSheetTypeSheetFormatPr();
                pr.defaultRowHeight = (value);
                pr.customHeight = (true);
            }
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
            return Workbook.GetCellStyleAt(idx == -1 ? 0 : idx);
        }

        /**
         * Whether the text is displayed in right-to-left mode in the window
         *
         * @return whether the text is displayed in right-to-left mode in the window
         */
        public bool RightToLeft
        {
            get
            {
                CT_SheetView view = GetDefaultSheetView();
                return view == null ? false : view.rightToLeft;
            }
            set
            {
                CT_SheetView view = GetDefaultSheetView();
                view.rightToLeft = (value);
            }
        }

        /**
         * Get whether to display the guts or not,
         * default value is true
         *
         * @return bool - guts or no guts
         */
        public bool DisplayGuts
        {
            get
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_OutlinePr outlinePr = sheetPr.outlinePr == null ? new CT_OutlinePr() : sheetPr.outlinePr;
                return outlinePr.showOutlineSymbols;
            }
            set
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_OutlinePr outlinePr = sheetPr.outlinePr == null ? sheetPr.AddNewOutlinePr() : sheetPr.outlinePr;
                outlinePr.showOutlineSymbols = (value);
            }
        }

        /**
         * Gets the flag indicating whether the window should show 0 (zero) in cells Containing zero value.
         * When false, cells with zero value appear blank instead of Showing the number zero.
         *
         * @return whether all zero values on the worksheet are displayed
         */
        public bool DisplayZeros
        {
            get
            {
                CT_SheetView view = GetDefaultSheetView();
                return view == null ? true : view.showZeros;
            }
            set
            {
                CT_SheetView view = GetSheetTypeSheetView();
                view.showZeros = (value);
            }
        }

        /**
         * Gets the first row on the sheet
         *
         * @return the number of the first logical row on the sheet, zero based
         */
        public int FirstRowNum
        {
            get
            {
                if (_rows.Count == 0)
                    return 0;
                else
                {
                    foreach (int key in _rows.Keys)
                    {
                        return key;
                    }
                    throw new ArgumentOutOfRangeException();
                }
            }
        }

        /**
         * Flag indicating whether the Fit to Page print option is enabled.
         *
         * @return <code>true</code>
         */
        public bool FitToPage
        {
            get
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_PageSetUpPr psSetup = (sheetPr == null || !sheetPr.IsSetPageSetUpPr()) ?
                        new CT_PageSetUpPr() : sheetPr.pageSetUpPr;
                return psSetup.fitToPage;
            }
            set
            {
                GetSheetTypePageSetUpPr().fitToPage = value;
            }
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
        public IFooter Footer
        {
            get
            {
                // The default footer is an odd footer
                return OddFooter;
            }
        }

        /**
         * Returns the default header for the sheet,
         *  creating one as needed.
         * You may also want to look at
         *  {@link #GetFirstHeader()},
         *  {@link #GetOddHeader()} and
         *  {@link #GetEvenHeader()}
         */
        public IHeader Header
        {
            get
            {
                // The default header is an odd header
                return OddHeader;
            }
        }

        /**
         * Returns the odd footer. Used on all pages unless
         *  other footers also present, when used on only
         *  odd pages.
         */
        public IFooter OddFooter
        {
            get
            {
                return new XSSFOddFooter(GetSheetTypeHeaderFooter());
            }
        }
        /**
         * Returns the even footer. Not there by default, but
         *  when Set, used on even pages.
         */
        public IFooter EvenFooter
        {
            get
            {
                return new XSSFEvenFooter(GetSheetTypeHeaderFooter());
            }
        }
        /**
         * Returns the first page footer. Not there by
         *  default, but when Set, used on the first page.
         */
        public IFooter FirstFooter
        {
            get
            {
                return new XSSFFirstFooter(GetSheetTypeHeaderFooter());
            }
        }

        /**
         * Returns the odd header. Used on all pages unless
         *  other headers also present, when used on only
         *  odd pages.
         */
        public IHeader OddHeader
        {
            get
            {
                return new XSSFOddHeader(GetSheetTypeHeaderFooter());
            }
        }
        /**
         * Returns the even header. Not there by default, but
         *  when Set, used on even pages.
         */
        public IHeader EvenHeader
        {
            get
            {
                return new XSSFEvenHeader(GetSheetTypeHeaderFooter());
            }
        }
        /**
         * Returns the first page header. Not there by
         *  default, but when Set, used on the first page.
         */
        public IHeader FirstHeader
        {
            get
            {
                return new XSSFFirstHeader(GetSheetTypeHeaderFooter());
            }
        }


        /**
         * Determine whether printed output for this sheet will be horizontally centered.
         */
        public bool HorizontallyCenter
        {
            get
            {
                CT_PrintOptions opts = worksheet.printOptions;
                return opts != null && opts.horizontalCentered;
            }
            set
            {
                CT_PrintOptions opts = worksheet.IsSetPrintOptions() ?
                    worksheet.printOptions : worksheet.AddNewPrintOptions();
                opts.horizontalCentered = (value);

            }
        }

        public int LastRowNum
        {
            get
            {
                return _rows.Count == 0 ? 0 : GetLastKey(_rows.Keys);
            }
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
        public double GetMargin(MarginType margin)
        {
            if (!worksheet.IsSetPageMargins()) return 0;

            CT_PageMargins pageMargins = worksheet.pageMargins;
            switch (margin)
            {
                case MarginType.LeftMargin:
                    return pageMargins.left;
                case MarginType.RightMargin:
                    return pageMargins.right;
                case MarginType.TopMargin:
                    return pageMargins.top;
                case MarginType.BottomMargin:
                    return pageMargins.bottom;
                case MarginType.HeaderMargin:
                    return pageMargins.header;
                case MarginType.FooterMargin:
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
        public void SetMargin(MarginType margin, double size)
        {
            CT_PageMargins pageMargins = worksheet.IsSetPageMargins() ?
                    worksheet.pageMargins : worksheet.AddNewPageMargins();
            switch (margin)
            {
                case MarginType.LeftMargin:
                    pageMargins.left = (size);
                    break;
                case MarginType.RightMargin:
                    pageMargins.right = (size);
                    break;
                case MarginType.TopMargin:
                    pageMargins.top = (size);
                    break;
                case MarginType.BottomMargin:
                    pageMargins.bottom = (size);
                    break;
                case MarginType.HeaderMargin:
                    pageMargins.header = (size);
                    break;
                case MarginType.FooterMargin:
                    pageMargins.footer = (size);
                    break;
                default:
                    throw new InvalidOperationException("Unknown margin constant:  " + margin);
            }
        }

        /**
         * @return the merged region at the specified index
         * @throws InvalidOperationException if this worksheet does not contain merged regions
         */
        public CellRangeAddress GetMergedRegion(int index)
        {
            CT_MergeCells ctMergeCells = worksheet.mergeCells;
            if (ctMergeCells == null) throw new InvalidOperationException("This worksheet does not contain merged regions");

            CT_MergeCell ctMergeCell = ctMergeCells.GetMergeCellArray(index);

            if (ctMergeCell == null) { return null; }

            String ref1 = ctMergeCell.@ref;
            return CellRangeAddress.ValueOf(ref1);
        }

        public CellRangeAddress GetMergedRegion(CellRangeAddress mergedRegion)
        {
            if (worksheet.mergeCells == null || worksheet.mergeCells.mergeCell == null)
                return null;
            foreach (CT_MergeCell mc in worksheet.mergeCells.mergeCell)
            {
                if (mc!=null && !string.IsNullOrEmpty(mc.@ref))
                {
                    CellRangeAddress range = CellRangeAddress.ValueOf(mc.@ref);
                    if (range.FirstColumn <= mergedRegion.FirstColumn
                     && range.LastColumn >= mergedRegion.LastColumn
                     && range.FirstRow <= mergedRegion.FirstRow
                     && range.LastRow >= mergedRegion.LastRow)
                    {
                        return range;
                    }
                }
            }
            return null;
        }

        /**
         * Returns the list of merged regions. If you want multiple regions, this is
         * faster than calling {@link #getMergedRegion(int)} each time.
         *
         * @return the list of merged regions
         * @throws InvalidOperationException if this worksheet does not contain merged regions
         */

        public List<CellRangeAddress> MergedRegions
        {
            get
            {
                List<CellRangeAddress> addresses = new List<CellRangeAddress>();
                CT_MergeCells ctMergeCells = worksheet.mergeCells;
                if (ctMergeCells == null) return addresses;

                foreach (CT_MergeCell ctMergeCell in ctMergeCells.mergeCell)
                {
                    String ref1 = ctMergeCell.@ref;
                    addresses.Add(CellRangeAddress.ValueOf(ref1));
                }
                return addresses;
            }

        }

        /**
         * Returns the number of merged regions defined in this worksheet
         *
         * @return number of merged regions in this worksheet
         */
        public int NumMergedRegions
        {
            get
            {
                CT_MergeCells ctMergeCells = worksheet.mergeCells;
                return ctMergeCells == null ? 0 : ctMergeCells.sizeOfMergeCellArray();
            }
        }

        public int NumHyperlinks
        {
            get
            {
                return hyperlinks.Count;
            }
        }

        /**
         * Returns the information regarding the currently configured pane (split or freeze).
         *
         * @return null if no pane configured, or the pane information.
         */
        public PaneInformation PaneInformation
        {
            get
            {
                CT_Pane pane = GetDefaultSheetView().pane;
                // no pane configured
                if (pane == null) return null;

                CellReference cellRef = pane.IsSetTopLeftCell() ?
                    new CellReference(pane.topLeftCell) : null;
                return new PaneInformation((short)pane.xSplit, (short)pane.ySplit,
                        (cellRef == null ? (short)0 : (short)cellRef.Row), (cellRef == null ? (short)0 : (short)cellRef.Col),
                        //in java the frist enum value is 1,but 0 in c#
                        (byte)(pane.activePane /*- 1*/), pane.state == ST_PaneState.frozen);
            }
        }

        /**
         * Returns the number of phsyically defined rows (NOT the number of rows in the sheet)
         *
         * @return the number of phsyically defined rows
         */
        public int PhysicalNumberOfRows
        {
            get
            {
                return _rows.Count;
            }
        }

        /**
         * Gets the print Setup object.
         *
         * @return The user model for the print Setup object.
         */
        public IPrintSetup PrintSetup
        {
            get
            {
                return new XSSFPrintSetup(worksheet);
            }
        }

        /**
         * Answer whether protection is enabled or disabled
         *
         * @return true => protection enabled; false => protection disabled
         */
        public bool Protect
        {
            get
            {
                return IsSheetLocked;
            }
        }

        /**
         * Enables sheet protection and Sets the password for the sheet.
         * Also Sets some attributes on the {@link CT_SheetProtection} that correspond to
         * the default values used by Excel
         * 
         * @param password to set for protection. Pass <code>null</code> to remove protection
         */
        public void ProtectSheet(String password)
        {

            if (password != null)
            {
                CT_SheetProtection sheetProtection = worksheet.AddNewSheetProtection();
                SetSheetPassword(password, null); // defaults to xor password
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
         * Sets the sheet password. 
         * 
         * @param password if null, the password will be removed
         * @param hashAlgo if null, the password will be set as XOR password (Excel 2010 and earlier)
         *  otherwise the given algorithm is used for calculating the hash password (Excel 2013)
         */
        public void SetSheetPassword(String password, HashAlgorithm hashAlgo)
        {
            if (password == null && !IsSheetProtectionEnabled())
            {
                return;
            }
            XSSFPasswordHelper.SetPassword(SafeGetProtectionField(), password, hashAlgo, null);
        }

        /**
         * Validate the password against the stored hash, the hashing method will be determined
         *  by the existing password attributes
         * @return true, if the hashes match (... though original password may differ ...)
         */
        public bool ValidateSheetPassword(String password)
        {
            if (!IsSheetProtectionEnabled())
            {
                return (password == null);
            }
            return XSSFPasswordHelper.ValidatePassword(SafeGetProtectionField(), password, null);
        }

        /**
         * Returns the logical row ( 0-based).  If you ask for a row that is not
         * defined you get a null.  This is to say row 4 represents the fifth row on a sheet.
         *
         * @param rownum  row to get
         * @return <code>XSSFRow</code> representing the rownumber or <code>null</code> if its not defined on the sheet
         */
        public IRow GetRow(int rownum)
        {
            if (_rows.ContainsKey(rownum))
                return _rows[rownum];
            return null;
        }


        /**
         * returns all rows between startRow and endRow, inclusive.
         * Rows between startRow and endRow that haven't been created are not included
         * in result unless createRowIfMissing is true
         *
         * @param startRow the first row number in this sheet to return
         * @param endRow the last row number in this sheet to return
         * @param createRowIfMissing
         * @return
         * @throws IllegalArgumentException if startRowNum and endRowNum are not in ascending order
         */
        private List<XSSFRow> GetRows(int startRowNum, int endRowNum, bool createRowIfMissing)
        {
            if (startRowNum > endRowNum)
            {
                throw new ArgumentException("getRows: startRowNum must be less than or equal to endRowNum");
            }
            List<XSSFRow> rows = new List<XSSFRow>();
            if (createRowIfMissing)
            {
                for (int i = startRowNum; i <= endRowNum; i++)
                {
                    XSSFRow row = GetRow(i) as XSSFRow;
                    if (row == null)
                    {
                        row = CreateRow(i) as XSSFRow;
                    }
                    rows.Add(row);
                }
            }
            else
            {
                //rows.addAll(_rows.subMap(startRowNum, endRowNum + 1).values());
                rows.AddRange(_rows.SkipWhile(x => x.Key < startRowNum)
                    .TakeWhile(x => x.Key < endRowNum + 1)
                    .Select(x => x.Value));
            }
            return rows;
        }

        /**
         * Horizontal page break information used for print layout view, page layout view, drawing print breaks in normal
         *  view, and for printing the worksheet.
         *
         * @return row indexes of all the horizontal page breaks, never <code>null</code>
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public int[] RowBreaks
        {
            get
            {
                if (!worksheet.IsSetRowBreaks() || worksheet.rowBreaks.sizeOfBrkArray() == 0)
                {
                    return new int[0];
                }

                List<CT_Break> brkArray = worksheet.rowBreaks.brk;
                int[] breaks = new int[brkArray.Count];
                for (int i = 0; i < brkArray.Count; i++)
                {
                    CT_Break brk = brkArray[i];
                    breaks[i] = (int)brk.id - 1;
                }
                return breaks;
            }
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
        public bool RowSumsBelow
        {
            get
            {
                CT_SheetPr sheetPr = worksheet.sheetPr;
                CT_OutlinePr outlinePr = (sheetPr != null && sheetPr.IsSetOutlinePr())
                        ? sheetPr.outlinePr : null;
                return outlinePr == null || outlinePr.summaryBelow;
            }
            set
            {
                ensureOutlinePr().summaryBelow = (value);
            }
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
        public bool RowSumsRight
        {
            get
            {
                CT_SheetPr sheetPr = worksheet.sheetPr;
                CT_OutlinePr outlinePr = (sheetPr != null && sheetPr.IsSetOutlinePr())
                        ? sheetPr.outlinePr : new CT_OutlinePr();
                return outlinePr.summaryRight;
            }
            set
            {
                ensureOutlinePr().summaryRight = (value);
            }
        }

        /**
         * Ensure CT_Worksheet.CT_SheetPr.CT_OutlinePr
         */
        private CT_OutlinePr ensureOutlinePr()
        {
            CT_SheetPr sheetPr = worksheet.IsSetSheetPr() ? worksheet.sheetPr : worksheet.AddNewSheetPr();
            return sheetPr.IsSetOutlinePr() ?
                sheetPr.outlinePr : sheetPr.AddNewOutlinePr();

        }

        /// <summary>
        /// A flag indicating whether scenarios are locked when the sheet is protected.
        /// </summary>
        public bool ScenarioProtect
        {
            get
            {
                return worksheet.IsSetSheetProtection()
                    && (bool)worksheet.sheetProtection.scenarios;
            }
        }
        public short LeftCol
        {
            get
            {
                String cellRef = GetPane().topLeftCell;
                if (cellRef == null)
                    return 0;
                CellReference cellReference = new CellReference(cellRef);
                return cellReference.Col;
            }
            set
            {
                throw new NotImplementedException();
            }
        }
        /// <summary>
        /// The top row in the visible view when the sheet is first viewed after opening it in a viewer
        /// </summary>
        public short TopRow
        {
            get
            {
                String cellRef = GetSheetTypeSheetView().topLeftCell;
                if (cellRef == null)
                    return 0;
                CellReference cellReference = new CellReference(cellRef);
                return (short)cellReference.Row;
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        /**
         * Determine whether printed output for this sheet will be vertically centered.
         *
         * @return whether printed output for this sheet will be vertically centered.
         */
        public bool VerticallyCenter
        {
            get
            {
                CT_PrintOptions opts = worksheet.printOptions;
                return opts != null && opts.verticalCentered;
            }
            set
            {
                CT_PrintOptions opts = worksheet.IsSetPrintOptions() ?
            worksheet.printOptions : worksheet.AddNewPrintOptions();
                opts.verticalCentered = (value);

            }
        }

        /**
         * Group between (0 based) columns
         */
        public void GroupColumn(int fromColumn, int toColumn)
        {
            GroupColumn1Based(fromColumn + 1, toColumn + 1);
        }
        private void GroupColumn1Based(int fromColumn, int toColumn)
        {
            CT_Cols ctCols = worksheet.GetColsArray(0);
            CT_Col ctCol = new CT_Col();

            // copy attributes, as they might be removed by merging with the new column
            // TODO: check if this fix is really necessary or if the sweeping algorithm
            // in addCleanColIntoCols needs to be adapted ...
            CT_Col fixCol_before = this.columnHelper.GetColumn1Based(toColumn, false);
            if (fixCol_before != null)
            {
                fixCol_before = (CT_Col)fixCol_before.Copy();
            }

            ctCol.min = (uint)fromColumn;
            ctCol.max = (uint)toColumn;
            this.columnHelper.AddCleanColIntoCols(ctCols, ctCol);

            CT_Col fixCol_after = this.columnHelper.GetColumn1Based(toColumn, false);
            if (fixCol_before != null && fixCol_after != null)
            {
                this.columnHelper.SetColumnAttributes(fixCol_before, fixCol_after);
            }

            for (int index = fromColumn; index <= toColumn; index++)
            {
                CT_Col col = columnHelper.GetColumn1Based(index, false);
                //col must exist
                short outlineLevel = col.outlineLevel;
                col.outlineLevel = (byte)(outlineLevel + 1);
                index = (int)col.max;
            }
            worksheet.SetColsArray(0, ctCols);
            SetSheetFormatPrOutlineLevelCol();
        }
        /**
         * Do not leave the width attribute undefined (see #52186).
         */
        private void SetColWidthAttribute(CT_Cols ctCols)
        {
            foreach (CT_Col col in ctCols.GetColList())
            {
                if (!col.IsSetWidth())
                {
                    col.width = (DefaultColumnWidth);
                    col.customWidth = (false);
                }
            }
        }
        /**
         * Tie a range of cell toGether so that they can be collapsed or expanded
         *
         * @param fromRow   start row (0-based)
         * @param toRow     end row (0-based)
         */
        public void GroupRow(int fromRow, int toRow)
        {
            for (int i = fromRow; i <= toRow; i++)
            {
                XSSFRow xrow = (XSSFRow)GetRow(i);
                if (xrow == null)
                {
                    xrow = (XSSFRow)CreateRow(i);
                }
                CT_Row ctrow = xrow.GetCTRow();
                short outlineLevel = ctrow.outlineLevel;
                ctrow.outlineLevel = ((byte)(outlineLevel + 1));
            }
            SetSheetFormatPrOutlineLevelRow();
        }

        private short GetMaxOutlineLevelRows()
        {
            short outlineLevel = 0;
            foreach (XSSFRow xrow in _rows.Values)
            {
                outlineLevel = xrow.GetCTRow().outlineLevel > outlineLevel ? xrow.GetCTRow().outlineLevel : outlineLevel;
            }
            return outlineLevel;
        }


        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        [Obsolete]
        private short GetMaxOutlineLevelCols()
        {
            CT_Cols ctCols = worksheet.GetColsArray(0);
            short outlineLevel = 0;
            foreach (CT_Col col in ctCols.GetColList())
            {
                outlineLevel = col.outlineLevel > outlineLevel ? col.outlineLevel : outlineLevel;
            }
            return outlineLevel;
        }

        /**
         * Determines if there is a page break at the indicated column
         */
        public bool IsColumnBroken(int column)
        {
            int[] colBreaks = ColumnBreaks;
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
            return col != null && (bool)col.hidden;
        }

        /**
         * Gets the flag indicating whether this sheet should display formulas.
         *
         * @return <code>true</code> if this sheet should display formulas.
         */
        public bool DisplayFormulas
        {
            get
            {
                return GetSheetTypeSheetView().showFormulas;
            }
            set
            {
                GetSheetTypeSheetView().showFormulas = value;
            }
        }

        /**
         * Gets the flag indicating whether this sheet displays the lines
         * between rows and columns to make editing and Reading easier.
         *
         * @return <code>true</code> if this sheet displays gridlines.
         * @see #isPrintGridlines() to check if printing of gridlines is turned on or off
         */
        public bool DisplayGridlines
        {
            get
            {
                return GetSheetTypeSheetView().showGridLines;
            }
            set
            {
                GetSheetTypeSheetView().showGridLines = value;
            }
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
        public bool DisplayRowColHeadings
        {
            get
            {
                return GetSheetTypeSheetView().showRowColHeaders;
            }
            set
            {
                GetSheetTypeSheetView().showRowColHeaders = value;
            }
        }


        /**
         * Returns whether gridlines are printed.
         *
         * @return whether gridlines are printed
         */
        public bool IsPrintGridlines
        {
            get
            {
                CT_PrintOptions opts = worksheet.printOptions;
                return opts != null && opts.gridLines;
            }
            set
            {
                CT_PrintOptions opts = worksheet.IsSetPrintOptions() ?
            worksheet.printOptions : worksheet.AddNewPrintOptions();
                opts.gridLines = (value);
            }
        }


        /**
         * Returns whether row and column headings are printed.
         *
         * @return whether row and column headings are printed
         */
        public bool IsPrintRowAndColumnHeadings
        {
            get
            {
                CT_PrintOptions opts = worksheet.printOptions;
                return opts != null && opts.headings;
            }
            set
            {
                CT_PrintOptions opts = worksheet.IsSetPrintOptions() ?
                    worksheet.printOptions : worksheet.AddNewPrintOptions();
                opts.headings = value;
            }
        }

        /**
         * Tests if there is a page break at the indicated row
         *
         * @param row index of the row to test
         * @return <code>true</code> if there is a page break at the indicated row
         */
        public bool IsRowBroken(int row)
        {
            int[] rowBreaks = RowBreaks;
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

            CT_PageBreak pgBreak = worksheet.IsSetRowBreaks() ? worksheet.rowBreaks : worksheet.AddNewRowBreaks();
            if (!IsRowBroken(row))
            {
                CT_Break brk = pgBreak.AddNewBrk();
                brk.id = (uint)row + 1; // this is id of the row element which is 1-based: <row r="1" ... >
                brk.man = (true);
                brk.max = (uint)(SpreadsheetVersion.EXCEL2007.LastColumnIndex); //end column of the break

                pgBreak.count = (uint)pgBreak.sizeOfBrkArray();
                pgBreak.manualBreakCount = (uint)pgBreak.sizeOfBrkArray();
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

            CT_PageBreak pgBreak = worksheet.colBreaks;
            List<CT_Break> brkArray = pgBreak.brk;
            for (int i = 0; i < brkArray.Count; i++)
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
            CT_MergeCells ctMergeCells = worksheet.mergeCells;

            int size = ctMergeCells.sizeOfMergeCellArray();
            CT_MergeCell[] mergeCellsArray = new CT_MergeCell[size - 1];
            for (int i = 0; i < size; i++)
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
                worksheet.UnsetMergeCells();
            }
        }

        /**
         * Removes a number of merged regions of cells (hence letting them free)
         * 
         * This method can be used to bulk-remove merged regions in a way
         * much faster than calling RemoveMergedRegion() for every single 
         * merged region.
         *
         * @param indices A Set of the regions to unmerge
         */
        public void RemoveMergedRegions(IList<int> indices)
        {
            if (!worksheet.IsSetMergeCells()) return;

            CT_MergeCells ctMergeCells = worksheet.mergeCells;
            //TODO: The following codes are not same as poi, re-do it?.
            int size = ctMergeCells.sizeOfMergeCellArray();
            List<CT_MergeCell> newMergeCells = new List<CT_MergeCell>(ctMergeCells.sizeOfMergeCellArray());

            for (int i = 0, d = 0; i < size; i++)
            {
                if (!indices.Contains(i))
                {
                    //newMergeCells[d] = ctMergeCells.GetMergeCellArray(i);
                    newMergeCells.Add(ctMergeCells.GetMergeCellArray(i));
                    d++;
                }
            }
            if (ListIsEmpty(newMergeCells))
            {
                worksheet.UnsetMergeCells();
            }
            else
            {
                ctMergeCells.SetMergeCellArray(newMergeCells.ToArray());
            }
        }
        private bool ListIsEmpty(List<CT_MergeCell> list)
        {
            foreach (CT_MergeCell mc in list)
            {
                if (mc != null)
                    return false;
            }
            return true;
        }

        /**
         * Remove a row from this sheet.  All cells Contained in the row are Removed as well
         *
         * @param row  the row to Remove.
         */
        public void RemoveRow(IRow row)
        {
            if (row.Sheet != this)
            {
                throw new ArgumentException("Specified row does not belong to this sheet");
            }
            // collect cells into a temporary array to avoid ConcurrentModificationException
            List<XSSFCell> cellsToDelete = new List<XSSFCell>();
            foreach (ICell cell in row) cellsToDelete.Add((XSSFCell)cell);

            foreach (XSSFCell cell in cellsToDelete) row.RemoveCell(cell);


            int idx = _rows.Count(p => p.Key < row.RowNum);// _rows.headMap(row.getRowNum()).size();
            _rows.Remove(row.RowNum);
            worksheet.sheetData.RemoveRow(idx); // Note that rows in worksheet.sheetData is 1-based.

            // also remove any comment located in that row
            if (sheetComments != null)
            {
                foreach (CellAddress ref1 in GetCellComments().Keys)
                {
                    if (ref1.Row == idx)
                    {
                        sheetComments.RemoveComment(ref1);
                    }
                }
            }
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
            CT_PageBreak pgBreak = worksheet.rowBreaks;
            List<CT_Break> brkArray = pgBreak.brk;
            for (int i = 0; i < brkArray.Count; i++)
            {
                if (brkArray[i].id == (row + 1))
                {
                    pgBreak.RemoveBrk(i);
                }
            }
        }

        /**
         * Whether Excel will be asked to recalculate all formulas when the
         *  workbook is opened.  
         */
        public bool ForceFormulaRecalculation
        {
            get
            {
                if (worksheet.IsSetSheetCalcPr())
                {
                    CT_SheetCalcPr calc = worksheet.sheetCalcPr;
                    return calc.fullCalcOnLoad;
                }
                return false;
            }
            set
            {
                CT_CalcPr calcPr = (Workbook as XSSFWorkbook).GetCTWorkbook().calcPr;
                if (worksheet.IsSetSheetCalcPr())
                {
                    // Change the current Setting
                    CT_SheetCalcPr calc = worksheet.sheetCalcPr;
                    calc.fullCalcOnLoad = value;
                }
                else if (value)
                {
                    // Add the Calc block and set it
                    CT_SheetCalcPr calc = worksheet.AddNewSheetCalcPr();
                    calc.fullCalcOnLoad = value;
                }

                if (value && calcPr != null && calcPr.calcMode == ST_CalcMode.manual)
                {
                    calcPr.calcMode = ST_CalcMode.auto;
                }
            }
        }

        /**
         * Flag indicating whether the sheet displays Automatic Page Breaks.
         *
         * @return <code>true</code> if the sheet displays Automatic Page Breaks.
         */
        public bool Autobreaks
        {
            get
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_PageSetUpPr psSetup = (sheetPr == null || !sheetPr.IsSetPageSetUpPr()) ?
                        new CT_PageSetUpPr() : sheetPr.pageSetUpPr;
                return psSetup.autoPageBreaks;
            }
            set
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_PageSetUpPr psSetup = sheetPr.IsSetPageSetUpPr() ?
                    sheetPr.pageSetUpPr : sheetPr.AddNewPageSetUpPr();
                psSetup.autoPageBreaks = (value);
            }
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
            if (!IsColumnBroken(column))
            {
                CT_PageBreak pgBreak = worksheet.IsSetColBreaks() ?
                    worksheet.colBreaks : worksheet.AddNewColBreaks();
                CT_Break brk = pgBreak.AddNewBrk();
                brk.id = (uint)column + 1;  // this is id of the row element which is 1-based: <row r="1" ... >
                brk.man = (true);
                brk.max = (uint)SpreadsheetVersion.EXCEL2007.LastRowIndex; //end row of the break

                pgBreak.count = (uint)pgBreak.sizeOfBrkArray();
                pgBreak.manualBreakCount = (uint)(pgBreak.sizeOfBrkArray());
            }
        }

        public void SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            if (collapsed)
            {
                CollapseColumn(columnNumber);
            }
            else
            {
                ExpandColumn(columnNumber);
            }
        }

        private void CollapseColumn(int columnNumber)
        {
            CT_Cols cols = worksheet.GetColsArray(0);
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
                    .outlineLevel, true);

            // write collapse field
            SetColumn(lastColMax + 1, null, 0, null, null, true);

        }

        private void SetColumn(int targetColumnIx, short? xfIndex, int? style,
                int? level, Boolean? hidden, Boolean? collapsed)
        {
            CT_Cols cols = worksheet.GetColsArray(0);
            CT_Col ci = null;
            int k = 0;
            for (k = 0; k < cols.sizeOfColArray(); k++)
            {
                CT_Col tci = cols.GetColArray(k);
                if (tci.min >= targetColumnIx
                        && tci.max <= targetColumnIx)
                {
                    ci = tci;
                    break;
                }
                if (tci.min > targetColumnIx)
                {
                    // call column infos after k are for later columns
                    break; // exit now so k will be the correct insert pos
                }
            }

            if (ci == null)
            {
                // okay so there ISN'T a column info record that covers this column
                // so lets create one!
                CT_Col nci = new CT_Col();
                nci.min = (uint)targetColumnIx;
                nci.max = (uint)targetColumnIx;
                UnsetCollapsed((bool)collapsed, nci);
                this.columnHelper.AddCleanColIntoCols(cols, nci);
                return;
            }

            bool styleChanged = style != null
            && ci.style != style;
            bool levelChanged = level != null
            && ci.outlineLevel != level;
            bool hiddenChanged = hidden != null
            && ci.hidden != hidden;
            bool collapsedChanged = collapsed != null
            && ci.collapsed != collapsed;
            bool columnChanged = levelChanged || hiddenChanged
            || collapsedChanged || styleChanged;
            if (!columnChanged)
            {
                // do nothing...nothing Changed.
                return;
            }

            if (ci.min == targetColumnIx && ci.max == targetColumnIx)
            {
                // ColumnInfo ci for a single column, the target column
                UnsetCollapsed((bool)collapsed, ci);
                return;
            }

            if (ci.min == targetColumnIx || ci.max == targetColumnIx)
            {
                // The target column is at either end of the multi-column ColumnInfo
                // ci
                // we'll just divide the info and create a new one
                if (ci.min == targetColumnIx)
                {
                    ci.min = (uint)(targetColumnIx + 1);
                }
                else
                {
                    ci.max = (uint)(targetColumnIx - 1);
                    k++; // adjust insert pos to insert after
                }
                CT_Col nci = columnHelper.CloneCol(cols, ci);
                nci.min = (uint)(targetColumnIx);
                UnsetCollapsed((bool)collapsed, nci);
                this.columnHelper.AddCleanColIntoCols(cols, nci);

            }
            else
            {
                // split to 3 records
                CT_Col ciStart = ci;
                CT_Col ciMid = columnHelper.CloneCol(cols, ci);
                CT_Col ciEnd = columnHelper.CloneCol(cols, ci);
                int lastcolumn = (int)ci.max;

                ciStart.max = (uint)(targetColumnIx - 1);

                ciMid.min = (uint)(targetColumnIx);
                ciMid.max = (uint)(targetColumnIx);
                UnsetCollapsed((bool)collapsed, ciMid);
                this.columnHelper.AddCleanColIntoCols(cols, ciMid);

                ciEnd.min = (uint)(targetColumnIx + 1);
                ciEnd.max = (uint)(lastcolumn);
                this.columnHelper.AddCleanColIntoCols(cols, ciEnd);
            }
        }

        private void UnsetCollapsed(bool collapsed, CT_Col ci)
        {
            if (collapsed)
            {
                ci.collapsed = (collapsed);
            }
            else
            {
                ci.UnsetCollapsed();
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
            CT_Cols cols = worksheet.GetColsArray(0);
            int idx = pIdx;
            CT_Col columnInfo = cols.GetColArray(idx);
            while (idx < cols.sizeOfColArray())
            {
                columnInfo.hidden = (hidden);
                if (idx + 1 < cols.sizeOfColArray())
                {
                    CT_Col nextColumnInfo = cols.GetColArray(idx + 1);

                    if (!IsAdjacentBefore(columnInfo, nextColumnInfo))
                    {
                        break;
                    }

                    if (nextColumnInfo.outlineLevel < level)
                    {
                        break;
                    }
                    columnInfo = nextColumnInfo;
                }
                idx++;
            }
            return (int)columnInfo.max;
        }

        private bool IsAdjacentBefore(CT_Col col, CT_Col other_col)
        {
            return (col.max == (other_col.min - 1));
        }

        private int FindStartOfColumnOutlineGroup(int pIdx)
        {
            // Find the start of the group.
            CT_Cols cols = worksheet.GetColsArray(0);
            CT_Col columnInfo = cols.GetColArray(pIdx);
            int level = columnInfo.outlineLevel;
            int idx = pIdx;
            while (idx != 0)
            {
                CT_Col prevColumnInfo = cols.GetColArray(idx - 1);
                if (!IsAdjacentBefore(prevColumnInfo, columnInfo))
                {
                    break;
                }
                if (prevColumnInfo.outlineLevel < level)
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
            CT_Cols cols = worksheet.GetColsArray(0);
            // Find the end of the group.
            CT_Col columnInfo = cols.GetColArray(colInfoIndex);
            int level = columnInfo.outlineLevel;
            int idx = colInfoIndex;
            while (idx < cols.sizeOfColArray() - 1)
            {
                CT_Col nextColumnInfo = cols.GetColArray(idx + 1);
                if (!IsAdjacentBefore(columnInfo, nextColumnInfo))
                {
                    break;
                }
                if (nextColumnInfo.outlineLevel < level)
                {
                    break;
                }
                idx++;
                columnInfo = nextColumnInfo;
            }
            return idx;
        }

        private void ExpandColumn(int columnIndex)
        {
            CT_Cols cols = worksheet.GetColsArray(0);
            CT_Col col = columnHelper.GetColumn(columnIndex, false);
            int colInfoIx = columnHelper.GetIndexOfColumn(cols, col);

            int idx = FindColInfoIdx((int)col.max, colInfoIx);
            if (idx == -1)
            {
                return;
            }

            // If it is already expanded do nothing.
            if (!IsColumnGroupCollapsed(idx))
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
            if (!IsColumnGroupHiddenByParent(idx))
            {
                int outlineLevel = columnInfo.outlineLevel;
                bool nestedGroup = false;
                for (int i = startIdx; i <= endIdx; i++)
                {
                    CT_Col ci = cols.GetColArray(i);
                    if (outlineLevel == ci.outlineLevel)
                    {
                        ci.UnsetHidden();
                        if (nestedGroup)
                        {
                            nestedGroup = false;
                            ci.collapsed = (true);
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
            SetColumn((int)columnInfo.max + 1, null, null, null,
                    false, false);
        }

        private bool IsColumnGroupHiddenByParent(int idx)
        {
            CT_Cols cols = worksheet.GetColsArray(0);
            // Look out outline details of end
            int endLevel = 0;
            bool endHidden = false;
            int endOfOutlineGroupIdx = FindEndOfColumnOutlineGroup(idx);
            if (endOfOutlineGroupIdx < cols.sizeOfColArray())
            {
                CT_Col nextInfo = cols.GetColArray(endOfOutlineGroupIdx + 1);
                if (IsAdjacentBefore(cols.GetColArray(endOfOutlineGroupIdx),
                        nextInfo))
                {
                    endLevel = nextInfo.outlineLevel;
                    endHidden = (bool)nextInfo.hidden;
                }
            }
            // Look out outline details of start
            int startLevel = 0;
            bool startHidden = false;
            int startOfOutlineGroupIdx = FindStartOfColumnOutlineGroup(idx);
            if (startOfOutlineGroupIdx > 0)
            {
                CT_Col prevInfo = cols.GetColArray(startOfOutlineGroupIdx - 1);

                if (IsAdjacentBefore(prevInfo, cols
                        .GetColArray(startOfOutlineGroupIdx)))
                {
                    startLevel = prevInfo.outlineLevel;
                    startHidden = (bool)prevInfo.hidden;
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
            CT_Cols cols = worksheet.GetColsArray(0);

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

                if (ci.min > fromColInfoIdx)
                {
                    break;
                }

            }
            return -1;
        }

        private bool ContainsColumn(CT_Col col, int columnIndex)
        {
            return col.min <= columnIndex && columnIndex <= col.max;
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
            CT_Cols cols = worksheet.GetColsArray(0);
            int endOfOutlineGroupIdx = FindEndOfColumnOutlineGroup(idx);
            int nextColInfoIx = endOfOutlineGroupIdx + 1;
            if (nextColInfoIx >= cols.sizeOfColArray())
            {
                return false;
            }
            CT_Col nextColInfo = cols.GetColArray(nextColInfoIx);

            CT_Col col = cols.GetColArray(endOfOutlineGroupIdx);
            if (!IsAdjacentBefore(col, nextColInfo))
            {
                return false;
            }

            return nextColInfo.collapsed;
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
         * </p>
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


        private CT_SheetView GetSheetTypeSheetView()
        {
            if (GetDefaultSheetView() == null)
            {
                GetSheetTypeSheetViews().SetSheetViewArray(0, new CT_SheetView());
            }
            return GetDefaultSheetView();
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
                CollapseRow(rowIndex);
            }
            else
            {
                ExpandRow(rowIndex);
            }
        }

        /**
         * @param rowIndex the zero based row index to collapse
         */
        private void CollapseRow(int rowIndex)
        {
            XSSFRow row = (XSSFRow)GetRow(rowIndex);
            if (row != null)
            {
                int startRow = FindStartOfRowOutlineGroup(rowIndex);

                // Hide all the columns until the end of the group
                int lastRow = WriteHidden(row, startRow, true);
                if (GetRow(lastRow) != null)
                {
                    ((XSSFRow)GetRow(lastRow)).GetCTRow().collapsed = (true);
                }
                else
                {
                    XSSFRow newRow = (XSSFRow)CreateRow(lastRow);
                    newRow.GetCTRow().collapsed = (true);
                }
            }
        }

        /**
         * @param rowIndex the zero based row index to find from
         */
        private int FindStartOfRowOutlineGroup(int rowIndex)
        {
            // Find the start of the group.
            int level = ((XSSFRow)GetRow(rowIndex)).GetCTRow().outlineLevel;
            int currentRow = rowIndex;
            while (GetRow(currentRow) != null)
            {
                if (((XSSFRow)GetRow(currentRow)).GetCTRow().outlineLevel < level)
                    return currentRow + 1;
                currentRow--;
            }
            return currentRow;
        }

        private int WriteHidden(XSSFRow xRow, int rowIndex, bool hidden)
        {
            int level = xRow.GetCTRow().outlineLevel;
            for (IEnumerator it = this.GetRowEnumerator(); it.MoveNext();)
            {
                xRow = (XSSFRow)it.Current;
                if (xRow.GetCTRow().outlineLevel >= level)
                {
                    xRow.GetCTRow().hidden = (hidden);
                    rowIndex++;
                }

            }
            return rowIndex;
        }

        /**
         * @param rowNumber the zero based row index to expand
         */
        private void ExpandRow(int rowNumber)
        {
            if (rowNumber == -1)
                return;
            XSSFRow row = (XSSFRow)GetRow(rowNumber);
            // If it is already expanded do nothing.
            if (!row.GetCTRow().IsSetHidden())
            {
                return;
            }
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
            if (!IsRowGroupHiddenByParent(rowNumber))
            {
                for (int i = startIdx; i < endIdx; i++)
                {
                    if (row.GetCTRow().outlineLevel == ((XSSFRow)GetRow(i)).GetCTRow()
                            .outlineLevel)
                    {
                        ((XSSFRow)GetRow(i)).GetCTRow().unsetHidden();
                    }
                    else if (!IsRowGroupCollapsed(i))
                    {
                        ((XSSFRow)GetRow(i)).GetCTRow().unsetHidden();
                    }
                }
            }
            // Write collapse field
            row = GetRow(endIdx) as XSSFRow;
            if (row != null)
            {
                CT_Row ctRow = row.GetCTRow();
                // This avoids an IndexOutOfBounds if multiple nested groups are collapsed/expanded
                if (ctRow.collapsed)
                {
                    ctRow.UnsetCollapsed();
                }
            }
        }

        /**
         * @param row the zero based row index to find from
         */
        public int FindEndOfRowOutlineGroup(int row)
        {
            int level = ((XSSFRow)GetRow(row)).GetCTRow().outlineLevel;
            int currentRow;
            int lastRowNum = LastRowNum;
            for (currentRow = row; currentRow < lastRowNum; currentRow++)
            {
                if (GetRow(currentRow) == null
                        || ((XSSFRow)GetRow(currentRow)).GetCTRow().outlineLevel < level)
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
                endLevel = ((XSSFRow)GetRow(endOfOutlineGroupIdx)).GetCTRow().outlineLevel;
                endHidden = (bool)((XSSFRow)GetRow(endOfOutlineGroupIdx)).GetCTRow().hidden;
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
                startLevel = ((XSSFRow)GetRow(startOfOutlineGroupIdx)).GetCTRow()
                .outlineLevel;
                startHidden = (bool)((XSSFRow)GetRow(startOfOutlineGroupIdx)).GetCTRow()
                .hidden;
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
            return (bool)((XSSFRow)GetRow(collapseRow)).GetCTRow().collapsed;
        }

        /**
         * Sets the zoom magnification for the sheet.  The zoom is expressed as a
         * fraction.  For example to express a zoom of 75% use 3 for the numerator
         * and 4 for the denominator.
         *
         * @param numerator     The numerator for the zoom magnification.
         * @param denominator   The denominator for the zoom magnification.
         * @see #SetZoom(int)
         */
        [Obsolete("deprecated 2015-11-23 (circa POI 3.14beta1). Use {@link #setZoom(int)} instead.")]
        public void SetZoom(int numerator, int denominator)
        {
            int zoom = 100 * numerator / denominator;
            SetZoom(zoom);
        }

        /**
         * Window zoom magnification for current view representing percent values.
         * Valid values range from 10 to 400. Horizontal &amp; Vertical scale toGether.
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
            if (scale < 10 || scale > 400)
                throw new ArgumentException("Valid scale values range from 10 to 400");
            GetSheetTypeSheetView().zoomScale = (uint)scale;
        }

        /**
         * copyRows rows from srcRows to this sheet starting at destStartRow
         *
         * Additionally copies merged regions that are completely defined in these
         * rows (ie. merged 2 cells on a row to be shifted).
         * @param srcRows the rows to copy. Formulas will be offset by the difference
         * in the row number of the first row in srcRows and destStartRow (even if srcRows
         * are from a different sheet).
         * @param destStartRow the row in this sheet to paste the first row of srcRows
         * the remainder of srcRows will be pasted below destStartRow per the cell copy policy
         * @param policy is the cell copy policy, which can be used to merge the source and destination
         * when the source is blank, copy styles only, paste as value, etc
         */

        public void CopyRows(List<XSSFRow> srcRows, int destStartRow, CellCopyPolicy policy)
        {
            if (srcRows == null || srcRows.Count == 0)
            {
                throw new ArgumentException("No rows to copy");
            }
            IRow srcStartRow = srcRows[0];
            IRow srcEndRow = srcRows[srcRows.Count - 1];

            if (srcStartRow == null)
            {
                throw new ArgumentException("copyRows: First row cannot be null");
            }

            int srcStartRowNum = srcStartRow.RowNum;
            int srcEndRowNum = srcEndRow.RowNum;

            // check row numbers to make sure they are continuous and increasing (monotonic)
            // and srcRows does not contain null rows
            int size = srcRows.Count;
            for (int index = 1; index < size; index++)
            {
                IRow curRow = srcRows[(index)];
                if (curRow == null)
                {
                    throw new ArgumentException("srcRows may not contain null rows. Found null row at index " + index + ".");
                    //} else if (curRow.RowNum != prevRow.RowNum + 1) {
                    //    throw new IllegalArgumentException("srcRows must contain continuously increasing row numbers. " +
                    //            "Got srcRows[" + (index-1) + "]=Row " + prevRow.RowNum + ", srcRows[" + index + "]=Row " + curRow.RowNum + ".");
                    // FIXME: assumes row objects belong to non-null sheets and sheets belong to non-null workbooks.
                }
                else if (srcStartRow.Sheet.Workbook != curRow.Sheet.Workbook)
                {
                    throw new ArgumentException("All rows in srcRows must belong to the same sheet in the same workbook." +
                            "Expected all rows from same workbook (" + srcStartRow.Sheet.Workbook + "). " +
                            "Got srcRows[" + index + "] from different workbook (" + curRow.Sheet.Workbook + ").");
                }
                else if (srcStartRow.Sheet != curRow.Sheet)
                {
                    throw new ArgumentException("All rows in srcRows must belong to the same sheet. " +
                            "Expected all rows from " + srcStartRow.Sheet.SheetName + ". " +
                            "Got srcRows[" + index + "] from " + curRow.Sheet.SheetName);
                }
            }

            // FIXME: is special behavior needed if srcRows and destRows belong to the same sheets and the regions overlap?

            CellCopyPolicy options = new CellCopyPolicy(policy);
            // avoid O(N^2) performance scanning through all regions for each row
            // merged regions will be copied after all the rows have been copied
            options.IsCopyMergedRegions = (false);

            // FIXME: if srcRows contains gaps or null values, clear out those rows that will be overwritten
            // how will this work with merging (copy just values, leave cell styles in place?)

            int r = destStartRow;
            foreach (IRow srcRow in srcRows)
            {
                int destRowNum;
                if (policy.IsCondenseRows)
                {
                    destRowNum = r++;
                }
                else
                {
                    int shift = (srcRow.RowNum - srcStartRowNum);
                    destRowNum = destStartRow + shift;
                }
                //removeRow(destRowNum); //this probably clears all external formula references to destRow, causing unwanted #REF! errors
                XSSFRow destRow = CreateRow(destRowNum) as XSSFRow;
                destRow.CopyRowFrom(srcRow, options);
            }

            // ======================
            // Only do additional copy operations here that cannot be done with Row.copyFromRow(Row, options)
            // reasons: operation needs to interact with multiple rows or sheets

            // Copy merged regions that are contained within the copy region
            if (policy.IsCopyMergedRegions)
            {
                // FIXME: is this something that rowShifter could be doing?
                int shift = destStartRow - srcStartRowNum;
                foreach (CellRangeAddress srcRegion in srcStartRow.Sheet.MergedRegions)
                {
                    if (srcStartRowNum <= srcRegion.FirstRow && srcRegion.LastRow <= srcEndRowNum)
                    {
                        // srcRegion is fully inside the copied rows
                        CellRangeAddress destRegion = srcRegion.Copy();
                        destRegion.FirstRow = (destRegion.FirstRow + shift);
                        destRegion.LastRow = (destRegion.LastRow + shift);
                        AddMergedRegion(destRegion);
                    }
                }
            }
        }

        /**
         * Copies rows between srcStartRow and srcEndRow to the same sheet, starting at destStartRow
         * Convenience function for {@link #copyRows(List, int, CellCopyPolicy)}
         * 
         * Equivalent to copyRows(getRows(srcStartRow, srcEndRow, false), destStartRow, cellCopyPolicy)
         * 
         * @param srcStartRow the index of the first row to copy the cells from in this sheet
         * @param srcEndRow the index of the last row to copy the cells from in this sheet
         * @param destStartRow the index of the first row to copy the cells to in this sheet
         * @param cellCopyPolicy the policy to use to determine how cells are copied
         */

        public void CopyRows(int srcStartRow, int srcEndRow, int destStartRow, CellCopyPolicy cellCopyPolicy)
        {
            List<XSSFRow> srcRows = GetRows(srcStartRow, srcEndRow, false); //FIXME: should be false, no need to create rows where src is only to copy them to dest
            CopyRows(srcRows, destStartRow, cellCopyPolicy);
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
         * rows (ie. merged 2 cells on a row to be Shifted).</p>
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
         * rows (ie. merged 2 cells on a row to be Shifted).</p>
         * 
         * @param startRow the row to start Shifting
         * @param endRow the row to end Shifting
         * @param n the number of rows to shift
         * @param copyRowHeight whether to copy the row height during the shift
         * @param reSetOriginalRowHeight whether to set the original row's height to the default
         */
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight)
        {
            XSSFVMLDrawing vml = GetVMLDrawing(false);
            List<int> rowsToRemove = new List<int>();
            List<CellAddress> commentsToRemove = new List<CellAddress>();
            List<CT_Row> ctRowsToRemove = new List<CT_Row>();
            // first remove all rows which will be overwritten
            foreach (KeyValuePair<int, XSSFRow> rowDict in _rows)
            {
                XSSFRow row = rowDict.Value;
                int rownum = row.RowNum;

                // check if we should remove this row as it will be overwritten by the data later
                if (ShouldRemoveRow(startRow, endRow, n, rownum))
                {
                    // remove row from worksheet.GetSheetData row array
                    //int idx = _rows.headMap(row.getRowNum()).size();
                    int idx = _rows.IndexOfValue(row);
                    //worksheet.sheetData.RemoveRow(idx);
                    ctRowsToRemove.Add(worksheet.sheetData.GetRowArray(idx));

                    // remove row from _rows
                    rowsToRemove.Add(rowDict.Key);

                    commentsToRemove.Clear();
                    // FIXME: (performance optimization) this should be moved outside the for-loop so that comments only needs to be iterated over once.
                    // also remove any comments associated with this row
                    if (sheetComments != null)
                    {
                        CT_CommentList lst = sheetComments.GetCTComments().commentList;
                        foreach (CT_Comment comment in lst.comment)
                        {
                            String strRef = comment.@ref;
                            CellAddress ref1 = new CellAddress(strRef);

                            // is this comment part of the current row?
                            if (ref1.Row == rownum)
                            {
                                //sheetComments.RemoveComment(strRef);
                                //vml.RemoveCommentShape(ref1.Row, ref1.Col);
                                commentsToRemove.Add(ref1);
                            }
                        }
                    }
                    foreach (CellAddress ref1 in commentsToRemove)
                    {
                        sheetComments.RemoveComment(ref1);
                        vml.RemoveCommentShape(ref1.Row, ref1.Column);
                    }

                    // FIXME: (performance optimization) this should be moved outside the for-loop so that hyperlinks only needs to be iterated over once.
                    // also remove any hyperlinks associated with this row
                    if (hyperlinks != null)
                    {
                        foreach (XSSFHyperlink link in new List<XSSFHyperlink>(hyperlinks))
                        {
                            CellReference ref1 = new CellReference(link.CellRef);
                            if (ref1.Row == rownum)
                            {
                                hyperlinks.Remove(link);
                            }
                        }
                    }
                }
            }
            foreach (int rowKey in rowsToRemove)
            {
                _rows.Remove(rowKey);
            }
            worksheet.sheetData.RemoveRows(ctRowsToRemove);
            // then do the actual moving and also adjust comments/rowHeight
            // we need to sort it in a way so the Shifting does not mess up the structures, 
            // i.e. when Shifting down, start from down and go up, when Shifting up, vice-versa
            SortedDictionary<XSSFComment, int> commentsToShift = new SortedDictionary<XSSFComment, int>(new ShiftCommentComparator(n));

            foreach (KeyValuePair<int, XSSFRow> rowDict in _rows)
            {
                XSSFRow row = rowDict.Value;
                int rownum = row.RowNum;

                if (sheetComments != null)
                {
                    // calculate the new rownum
                    int newrownum = ShiftedRowNum(startRow, endRow, n, rownum);

                    // is there a change necessary for the current row?
                    if (newrownum != rownum)
                    {
                        CT_CommentList lst = sheetComments.GetCTComments().commentList;
                        foreach (CT_Comment comment in lst.comment)
                        {
                            String oldRef = comment.@ref;
                            CellReference ref1 = new CellReference(oldRef);

                            // is this comment part of the current row?
                            if (ref1.Row == rownum)
                            {
                                XSSFComment xssfComment = new XSSFComment(sheetComments, comment,
                                        vml == null ? null : vml.FindCommentShape(rownum, ref1.Col));

                                // we should not perform the Shifting right here as we would then find
                                // already Shifted comments and would shift them again...
                                if (commentsToShift.ContainsKey(xssfComment))
                                    commentsToShift[xssfComment] = newrownum;
                                else
                                    commentsToShift.Add(xssfComment, newrownum);
                            }
                        }
                    }
                }

                if (rownum < startRow || rownum > endRow) continue;

                if (!copyRowHeight)
                {
                    row.Height = (/*setter*/(short)-1);
                }

                row.Shift(n);
            }

            // adjust all the affected comment-structures now
            // the Map is sorted and thus provides them in the order that we need here, 
            // i.e. from down to up if Shifting down, vice-versa otherwise
            foreach (KeyValuePair<XSSFComment, int> entry in commentsToShift)
            {
                entry.Key.Row = (/*setter*/entry.Value);
            }

            XSSFRowShifter rowShifter = new XSSFRowShifter(this);

            int sheetIndex = Workbook.GetSheetIndex(this);
            String sheetName = Workbook.GetSheetName(sheetIndex);
            FormulaShifter shifter = FormulaShifter.CreateForRowShift(
                                       sheetIndex, sheetName, startRow, endRow, n, SpreadsheetVersion.EXCEL2007);

            rowShifter.UpdateNamedRanges(shifter);
            rowShifter.UpdateFormulas(shifter);
            rowShifter.ShiftMergedRegions(startRow, endRow, n);
            rowShifter.UpdateConditionalFormatting(shifter);
            rowShifter.UpdateHyperlinks(shifter);

            //rebuild the _rows map
            Dictionary<int, XSSFRow> map = new Dictionary<int, XSSFRow>();
            foreach (XSSFRow r in _rows.Values)
            {
                map.Add(r.RowNum, r);
            }
            _rows.Clear();
            //_rows.putAll(map);
            foreach (KeyValuePair<int, XSSFRow> kv in map)
            {
                _rows.Add(kv.Key, kv.Value);
            }

            // Sort CTRows by index asc.
            // not found at poi 3.15
            if (worksheet.sheetData.row != null)
                worksheet.sheetData.row.Sort((row1, row2) => row1.r.CompareTo(row2.r));
        }
        private class ShiftCommentComparator : IComparer<XSSFComment>
        {
            private int shiftDir;
            public ShiftCommentComparator(int shiftDir)
            {
                this.shiftDir = shiftDir;
            }
            public int Compare(XSSFComment o1, XSSFComment o2)
            {
                int row1 = o1.Row;
                int row2 = o2.Row;

                if (row1 == row2)
                {
                    // ordering is not important when row is equal, but don't return zero to still 
                    // get multiple comments per row into the map
                    return o1.GetHashCode() - o2.GetHashCode();
                }

                // when Shifting down, sort higher row-values first
                if (shiftDir > 0)
                {
                    return row1 < row2 ? 1 : -1;
                }
                else
                {
                    // sort lower-row values first when Shifting up
                    return row1 > row2 ? 1 : -1;
                }
            }
        }
        private int ShiftedRowNum(int startRow, int endRow, int n, int rownum)
        {
            // no change if before any affected row
            if (rownum < startRow && (n > 0 || (startRow - rownum) > n))
            {
                return rownum;
            }

            // no change if After any affected row
            if (rownum > endRow && (n < 0 || (rownum - endRow) > n))
            {
                return rownum;
            }

            // row before and things are Moved up
            if (rownum < startRow)
            {
                // row is Moved down by the Shifting
                return rownum + (endRow - startRow);
            }

            // row is After and things are Moved down
            if (rownum > endRow)
            {
                // row is Moved up by the Shifting
                return rownum - (endRow - startRow);
            }

            // row is part of the Shifted block
            return rownum + n;
        }

        public void UngroupColumn(int fromColumn, int toColumn)
        {
            CT_Cols cols = worksheet.GetColsArray(0);
            for (int index = fromColumn; index <= toColumn; index++)
            {
                CT_Col col = columnHelper.GetColumn(index, false);
                if (col != null)
                {
                    short outlineLevel = col.outlineLevel;
                    col.outlineLevel = (byte)(outlineLevel - 1);
                    index = (int)col.max;

                    if (col.outlineLevel <= 0)
                    {
                        int colIndex = columnHelper.GetIndexOfColumn(cols, col);
                        worksheet.GetColsArray(0).RemoveCol(colIndex);
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
                XSSFRow xrow = (XSSFRow)GetRow(i);
                if (xrow != null)
                {
                    CT_Row ctrow = xrow.GetCTRow();
                    short outlinelevel = ctrow.outlineLevel;
                    ctrow.outlineLevel = (byte)(outlinelevel - 1);
                    //remove a row only if the row has no cell and if the outline level is 0
                    if (ctrow.outlineLevel == 0 && xrow.FirstCellNum == -1)
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
            GetSheetTypeSheetFormatPr().outlineLevelRow = (byte)maxLevelRow;
        }

        private void SetSheetFormatPrOutlineLevelCol()
        {
            short maxLevelCol = GetMaxOutlineLevelCols();
            GetSheetTypeSheetFormatPr().outlineLevelCol = (byte)maxLevelCol;
        }

        private CT_SheetViews GetSheetTypeSheetViews()
        {
            if (worksheet.sheetViews == null)
            {
                worksheet.sheetViews = new CT_SheetViews();
                worksheet.sheetViews.AddNewSheetView();
            }
            return worksheet.sheetViews;
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
        public bool IsSelected
        {
            get
            {
                CT_SheetView view = GetDefaultSheetView();
                return view != null && view.tabSelected;
            }
            set
            {
                CT_SheetViews views = GetSheetTypeSheetViews();
                foreach (CT_SheetView view in views.sheetView)
                {
                    view.tabSelected = (value);
                }
            }
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
         * Removes a hyperlink in the collection of hyperlinks on this sheet
         *
         * @param row row index
         * @param column column index
         */
        public void RemoveHyperlink(int row, int column)
        {
            // CTHyperlinks is regenerated from scratch when writing out the spreadsheet
            // so don't worry about maintaining hyperlinks and CTHyperlinks in parallel.
            // only maintain hyperlinks
            String ref1 = new CellReference(row, column).FormatAsString();
            for (int index = 0; index < hyperlinks.Count; index++)
            {
                XSSFHyperlink hyperlink = hyperlinks[index];
                if (hyperlink.CellRef.Equals(ref1))
                {
                    hyperlinks.RemoveAt(index);
                    return;
                }
            }
        }
        /**
         * Return location of the active cell, e.g. <code>A1</code>.
         *
         * @return the location of the active cell.
         */
        public CellAddress ActiveCell
        {
            get
            {
                String address = GetSheetTypeSelection().activeCell;
                if (address == null)
                {
                    return null;
                }
                return new CellAddress(address);
            }
            set
            {
                String ref1 = value.FormatAsString();
                CT_Selection ctsel = GetSheetTypeSelection();
                ctsel.activeCell = (ref1);
                ctsel.SetSqref(new string[] { ref1 });
            }
        }
        [Obsolete("deprecated 3.14beta2 (circa 2015-12-05). Use {@link #setActiveCell(CellAddress)} instead.")]
        public void SetActiveCell(string cellref)
        {
            CT_Selection ctsel = GetSheetTypeSelection();
            ctsel.activeCell = cellref;
            ctsel.SetSqref(new string[] { cellref });
        }

        //public void SetActiveCell(int row, int column)
        //{
        //    CellReference cellref = new CellReference(row, column);
        //    SetActiveCell(cellref.FormatAsString());
        //}
        /**
         * Does this sheet have any comments on it? We need to know,
         *  so we can decide about writing it to disk or not
         */
        public bool HasComments
        {
            get
            {
                if (sheetComments == null) { return false; }
                return (sheetComments.GetNumberOfComments() > 0);
            }
        }

        internal int NumberOfComments
        {
            get
            {
                if (sheetComments == null) { return 0; }
                return sheetComments.GetNumberOfComments();
            }
        }

        private CT_Selection GetSheetTypeSelection()
        {
            if (GetSheetTypeSheetView().SizeOfSelectionArray() == 0)
            {
                GetSheetTypeSheetView().InsertNewSelection(0);
            }

            return GetSheetTypeSheetView().GetSelectionArray(0);
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
        protected internal CommentsTable GetCommentsTable(bool create)
        {
            if (sheetComments == null && create)
            {
                // Try to create a comments table with the same number as
                //  the sheet has (i.e. sheet 1 -> comments 1)
                try
                {
                    sheetComments = (CommentsTable)CreateRelationship(
                          XSSFRelation.SHEET_COMMENTS, XSSFFactory.GetInstance(), (int)sheet.sheetId);
                }
                catch (PartAlreadyExistsException)
                {
                    // Technically a sheet doesn't need the same number as
                    //  it's comments, and clearly someone has already pinched
                    //  our number! Go for the next available one instead
                    sheetComments = (CommentsTable)CreateRelationship(
                          XSSFRelation.SHEET_COMMENTS, XSSFFactory.GetInstance(), -1);
                }
            }
            return sheetComments;
        }

        private CT_PageSetUpPr GetSheetTypePageSetUpPr()
        {
            CT_SheetPr sheetPr = GetSheetTypeSheetPr();
            return sheetPr.IsSetPageSetUpPr() ? sheetPr.pageSetUpPr : sheetPr.AddNewPageSetUpPr();
        }

        private static bool ShouldRemoveRow(int startRow, int endRow, int n, int rownum)
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
            if (GetDefaultSheetView().pane == null)
            {
                GetDefaultSheetView().AddNewPane();
            }
            return GetDefaultSheetView().pane;
        }

        /**
         * Return a master shared formula by index
         *
         * @param sid shared group index
         * @return a CT_CellFormula bean holding shared formula or <code>null</code> if not found
         */
        internal CT_CellFormula GetSharedFormula(int sid)
        {
            return sharedFormulas[sid];
        }

        internal void OnReadCell(XSSFCell cell)
        {
            //collect cells holding shared formulas
            CT_Cell ct = cell.GetCTCell();
            CT_CellFormula f = ct.f;
            if (f != null && f.t == ST_CellFormulaType.shared && f.isSetRef() && f.Value != null)
            {
                // save a detached  copy to avoid XmlValueDisconnectedException,
                // this may happen when the master cell of a shared formula is Changed
                CT_CellFormula sf = (CT_CellFormula)f.Copy();
                CellRangeAddress sfRef = CellRangeAddress.ValueOf(sf.@ref);
                CellReference cellRef = new CellReference(cell);
                // If the shared formula range preceeds the master cell then the preceding  part is discarded, e.g.
                // if the cell is E60 and the shared formula range is C60:M85 then the effective range is E60:M85
                // see more details in https://issues.apache.org/bugzilla/show_bug.cgi?id=51710
                if (cellRef.Col > sfRef.FirstColumn || cellRef.Row > sfRef.FirstRow)
                {
                    String effectiveRef = new CellRangeAddress(
                            Math.Max(cellRef.Row, sfRef.FirstRow), sfRef.LastRow,
                            Math.Max(cellRef.Col, sfRef.FirstColumn), sfRef.LastColumn).FormatAsString();
                    sf.@ref = (effectiveRef);
                }
                sharedFormulas[(int)f.si] = sf;
            }
            if (f != null && f.t == ST_CellFormulaType.array && f.@ref != null)
            {
                arrayFormulas.Add(CellRangeAddress.ValueOf(f.@ref));
            }
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            Write(out1);
            out1.Close();
        }

        internal virtual void Write(Stream stream)
        {
            bool setToNull = false;
            if (worksheet.sizeOfColsArray() == 1)
            {
                CT_Cols col = worksheet.GetColsArray(0);
                if (col.sizeOfColArray() == 0)
                {
                    setToNull = true;
                    // this is necessary so that we do not write an empty <cols/> item into the sheet-xml in the xlsx-file
                    // Excel complains about a corrupted file if this shows up there!
                    worksheet.SetColsArray(null);
                }
                else
                {
                    SetColWidthAttribute(col);
                }
            }


            // Now re-generate our CT_Hyperlinks, if needed
            if (hyperlinks.Count > 0)
            {
                if (worksheet.hyperlinks == null)
                {
                    worksheet.AddNewHyperlinks();
                }
                NPOI.OpenXmlFormats.Spreadsheet.CT_Hyperlink[] ctHls
                    = new NPOI.OpenXmlFormats.Spreadsheet.CT_Hyperlink[hyperlinks.Count];
                for (int i = 0; i < ctHls.Length; i++)
                {
                    // If our sheet has hyperlinks, have them add
                    //  any relationships that they might need
                    XSSFHyperlink hyperlink = hyperlinks[i];
                    hyperlink.GenerateRelationIfNeeded(GetPackagePart());
                    // Now grab their underling object
                    ctHls[i] = hyperlink.GetCTHyperlink();
                }
                worksheet.hyperlinks.SetHyperlinkArray(ctHls);
            }

            foreach (XSSFRow row in _rows.Values)
            {
                row.OnDocumentWrite();
            }

            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //xmlOptions.SetSaveSyntheticDocumentElement(new QName(CT_Worksheet.type.GetName().getNamespaceURI(), "worksheet"));
            Dictionary<String, String> map = new Dictionary<String, String>();
            map[ST_RelationshipId.NamespaceURI] = "r";
            //xmlOptions.SetSaveSuggestedPrefixes(map);

            new WorksheetDocument(worksheet).Save(stream);

            // Bug 52233: Ensure that we have a col-array even if write() removed it
            if (setToNull)
            {
                worksheet.AddNewCols();
            }
        }

        /**
         * @return true when Autofilters are locked and the sheet is protected.
         */
        public bool IsAutoFilterLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().autoFilter;
            }
        }

        /**
         * @return true when Deleting columns is locked and the sheet is protected.
         */
        public bool IsDeleteColumnsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().deleteColumns;
            }
        }

        /**
         * @return true when Deleting rows is locked and the sheet is protected.
         */
        public bool IsDeleteRowsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().deleteRows;
            }
        }

        /**
         * @return true when Formatting cells is locked and the sheet is protected.
         */
        public bool IsFormatCellsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().formatCells;
            }
        }

        /**
         * @return true when Formatting columns is locked and the sheet is protected.
         */
        public bool IsFormatColumnsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().formatColumns;
            }
        }

        /**
         * @return true when Formatting rows is locked and the sheet is protected.
         */
        public bool IsFormatRowsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().formatRows;
            }
        }

        /**
         * @return true when Inserting columns is locked and the sheet is protected.
         */
        public bool IsInsertColumnsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().insertColumns;
            }
        }

        /**
         * @return true when Inserting hyperlinks is locked and the sheet is protected.
         */
        public bool IsInsertHyperlinksLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().insertHyperlinks;
            }
        }

        /**
         * @return true when Inserting rows is locked and the sheet is protected.
         */
        public bool IsInsertRowsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().insertRows;
            }
        }

        /**
         * @return true when Pivot tables are locked and the sheet is protected.
         */
        public bool IsPivotTablesLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().pivotTables;
            }
        }

        /**
         * @return true when Sorting is locked and the sheet is protected.
         */
        public bool IsSortLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().sort;
            }
        }

        /**
         * @return true when Objects are locked and the sheet is protected.
         */
        public bool IsObjectsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().objects;
            }
        }

        /**
         * @return true when Scenarios are locked and the sheet is protected.
         */
        public bool IsScenariosLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().scenarios;
            }
        }

        /**
         * @return true when Selection of locked cells is locked and the sheet is protected.
         */
        public bool IsSelectLockedCellsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().selectLockedCells;
            }
        }

        /**
         * @return true when Selection of unlocked cells is locked and the sheet is protected.
         */
        public bool IsSelectUnlockedCellsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().selectUnlockedCells;
            }
        }

        /**
         * @return true when Sheet is Protected.
         */
        public bool IsSheetLocked
        {
            get
            {
                return worksheet.IsSetSheetProtection() && SafeGetProtectionField().sheet;
            }
        }

        /**
         * Enable sheet protection
         */
        public void EnableLocking()
        {
            SafeGetProtectionField().sheet = true;
        }

        /**
         * Disable sheet protection
         */
        public void DisableLocking()
        {
            SafeGetProtectionField().sheet = false;
        }

        /**
         * Enable or disable Autofilters locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockAutoFilter(bool enabled)
        {
            SafeGetProtectionField().autoFilter = enabled;
        }

        /**
         * Enable or disable Deleting columns locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockDeleteColumns(bool enabled)
        {
            SafeGetProtectionField().deleteColumns = enabled;
        }

        /**
         * Enable or disable Deleting rows locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockDeleteRows(bool enabled)
        {
            SafeGetProtectionField().deleteRows = enabled;
        }

        /**
         * Enable or disable Formatting cells locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockFormatCells(bool enabled)
        {
            SafeGetProtectionField().formatCells = enabled;
        }

        /**
         * Enable or disable Formatting columns locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockFormatColumns(bool enabled)
        {
            SafeGetProtectionField().formatColumns = enabled;
        }

        /**
         * Enable or disable Formatting rows locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockFormatRows(bool enabled)
        {
            SafeGetProtectionField().formatRows = enabled;
        }

        /**
         * Enable or disable Inserting columns locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockInsertColumns(bool enabled)
        {
            SafeGetProtectionField().insertColumns = enabled;
        }

        /**
         * Enable or disable Inserting hyperlinks locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockInsertHyperlinks(bool enabled)
        {
            SafeGetProtectionField().insertHyperlinks = enabled;
        }

        /**
         * Enable or disable Inserting rows locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockInsertRows(bool enabled)
        {
            SafeGetProtectionField().insertRows = enabled;
        }

        /**
         * Enable or disable Pivot Tables locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockPivotTables(bool enabled)
        {
            SafeGetProtectionField().pivotTables = enabled;
        }

        /**
         * Enable or disable Sort locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockSort(bool enabled)
        {
            SafeGetProtectionField().sort = enabled;
        }

        /**
         * Enable or disable Objects locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockObjects(bool enabled)
        {
            SafeGetProtectionField().objects = enabled;
        }

        /**
         * Enable or disable Scenarios locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockScenarios(bool enabled)
        {
            SafeGetProtectionField().scenarios = enabled;
        }

        /**
         * Enable or disable Selection of locked cells locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockSelectLockedCells(bool enabled)
        {
            SafeGetProtectionField().selectLockedCells = enabled;
        }

        /**
         * Enable or disable Selection of unlocked cells locking.
         * This does not modify sheet protection status.
         * To enforce this un-/locking, call {@link #disableLocking()} or {@link #enableLocking()}
         */
        public void LockSelectUnlockedCells(bool enabled)
        {
            SafeGetProtectionField().selectUnlockedCells = enabled;
        }

        private CT_SheetProtection SafeGetProtectionField()
        {
            if (!IsSheetProtectionEnabled())
            {
                return worksheet.AddNewSheetProtection();
            }
            return worksheet.sheetProtection;
        }

        /* package */
        bool IsSheetProtectionEnabled()
        {
            return (worksheet.IsSetSheetProtection());
        }


        /* namespace */
        internal bool IsCellInArrayFormulaContext(ICell cell)
        {
            foreach (CellRangeAddress range in arrayFormulas)
            {
                if (range.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return true;
                }
            }
            return false;
        }

        /* namespace */
        internal XSSFCell GetFirstCellInArrayFormula(ICell cell)
        {
            foreach (CellRangeAddress range in arrayFormulas)
            {
                if (range.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return (XSSFCell)GetRow(range.FirstRow).GetCell(range.FirstColumn);
                }
            }
            return null;
        }

        /**
         * Also Creates cells if they don't exist
         */
        private ICellRange<ICell> GetCellRange(CellRangeAddress range)
        {
            int firstRow = range.FirstRow;
            int firstColumn = range.FirstColumn;
            int lastRow = range.LastRow;
            int lastColumn = range.LastColumn;
            int height = lastRow - firstRow + 1;
            int width = lastColumn - firstColumn + 1;
            List<ICell> temp = new List<ICell>(height * width);
            for (int rowIn = firstRow; rowIn <= lastRow; rowIn++)
            {
                for (int colIn = firstColumn; colIn <= lastColumn; colIn++)
                {
                    IRow row = GetRow(rowIn);
                    if (row == null)
                    {
                        row = CreateRow(rowIn);
                    }
                    ICell cell = row.GetCell(colIn);
                    if (cell == null)
                    {
                        cell = row.CreateCell(colIn);
                    }
                    temp.Add(cell);
                }
            }
            return SSCellRange<ICell>.Create(firstRow, firstColumn, height, width, temp, typeof(ICell));
        }

        public ICellRange<ICell> SetArrayFormula(String formula, CellRangeAddress range)
        {

            ICellRange<ICell> cr = GetCellRange(range);

            ICell mainArrayFormulaCell = cr.TopLeftCell;
            ((XSSFCell)mainArrayFormulaCell).SetCellArrayFormula(formula, range);
            arrayFormulas.Add(range);
            return cr;
        }

        public ICellRange<ICell> RemoveArrayFormula(ICell cell)
        {
            if (cell.Sheet != this)
            {
                throw new ArgumentException("Specified cell does not belong to this sheet.");
            }
            foreach (CellRangeAddress range in arrayFormulas)
            {
                if (range.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    arrayFormulas.Remove(range);
                    ICellRange<ICell> cr = GetCellRange(range);
                    foreach (ICell c in cr)
                    {
                        c.SetCellType(CellType.Blank);
                    }
                    return cr;
                }
            }
            String ref1 = ((XSSFCell)cell).GetCTCell().r;
            throw new ArgumentException("Cell " + ref1 + " is not part of an array formula.");
        }


        public IDataValidationHelper GetDataValidationHelper()
        {
            return dataValidationHelper;
        }

        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public List<IDataValidation> GetDataValidations()
        {
            List<IDataValidation> xssfValidations = new List<IDataValidation>();
            CT_DataValidations dataValidations = this.worksheet.dataValidations;
            if (dataValidations != null && dataValidations.count > 0)
            {
                foreach (CT_DataValidation ctDataValidation in dataValidations.dataValidation)
                {
                    CellRangeAddressList addressList = new CellRangeAddressList();

                    String[] regions = ctDataValidation.sqref.Split(new char[] { ' ' });
                    for (int i = 0; i < regions.Length; i++)
                    {
                        if (regions[i].Length == 0)
                            continue;
                        String[] parts = regions[i].Split(new char[] { ':' });
                        CellReference begin = new CellReference(parts[0]);
                        CellReference end = parts.Length > 1 ? new CellReference(parts[1]) : begin;
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(begin.Row, end.Row, begin.Col, end.Col);
                        addressList.AddCellRangeAddress(cellRangeAddress);
                    }
                    XSSFDataValidation xssfDataValidation = new XSSFDataValidation(addressList, ctDataValidation);
                    xssfValidations.Add(xssfDataValidation);
                }
            }
            return xssfValidations;
        }

        public void AddValidationData(IDataValidation dataValidation)
        {
            XSSFDataValidation xssfDataValidation = (XSSFDataValidation)dataValidation;
            CT_DataValidations dataValidations = worksheet.dataValidations;
            if (dataValidations == null)
            {
                dataValidations = worksheet.AddNewDataValidations();
            }

            int currentCount = dataValidations.sizeOfDataValidationArray();
            CT_DataValidation newval = dataValidations.AddNewDataValidation();
            newval.Set(xssfDataValidation.GetCTDataValidation());
            dataValidations.count = (uint)currentCount + 1;

        }

        public IAutoFilter SetAutoFilter(CellRangeAddress range)
        {
            CT_AutoFilter af = worksheet.autoFilter;
            if (af == null) af = worksheet.AddNewAutoFilter();

            CellRangeAddress norm = new CellRangeAddress(range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn);
            String ref1 = norm.FormatAsString();
            af.@ref = (ref1);

            XSSFWorkbook wb = (XSSFWorkbook)Workbook;
            int sheetIndex = Workbook.GetSheetIndex(this);
            XSSFName name = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
            if (name == null)
            {
                name = wb.CreateBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
            }
            name.GetCTName().hidden = true;
            CellReference r1 = new CellReference(SheetName, range.FirstRow, range.FirstColumn, true, true);
            CellReference r2 = new CellReference(null, range.LastRow, range.LastColumn, true, true);
            String fmla = r1.FormatAsString() + ":" + r2.FormatAsString();
            name.RefersToFormula = fmla;

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

            CT_TableParts tblParts = worksheet.tableParts;
            CT_TablePart tbl = tblParts.AddNewTablePart();

            // Table numbers need to be unique in the file, not just
            //  unique within the sheet. Find the next one
            int tableNumber = GetPackagePart().Package.GetPartsByContentType(XSSFRelation.TABLE.ContentType).Count + 1;
            RelationPart rp = CreateRelationship(XSSFRelation.TABLE, XSSFFactory.GetInstance(), tableNumber, false);
            XSSFTable table = rp.DocumentPart as XSSFTable;
            tbl.id = rp.Relationship.Id;

            tables[tbl.id] = table;

            return table;
        }

        /**
         * Returns any tables associated with this Sheet
         */
        public List<XSSFTable> GetTables()
        {
            List<XSSFTable> tableList = new List<XSSFTable>(
                  tables.Values
            );
            return tableList;
        }

        public ISheetConditionalFormatting SheetConditionalFormatting
        {
            get
            {
                return new XSSFSheetConditionalFormatting(this);
            }
        }
        /**
         * Set background color of the sheet tab
         *
         * @param colorIndex  the indexed color to set, must be a constant from {@link IndexedColors}
         */
        [Obsolete("deprecated 3.15-beta2. Removed in 3.17. Use {@link #setTabColor(XSSFColor)}.")]
        public void SetTabColor(int colorIndex)
        {
            CT_SheetPr pr = worksheet.sheetPr;
            if (pr == null) pr = worksheet.AddNewSheetPr();
            NPOI.OpenXmlFormats.Spreadsheet.CT_Color color = new OpenXmlFormats.Spreadsheet.CT_Color();
            color.indexed = (uint)(colorIndex);
            pr.tabColor = (color);
        }

        /*
         * Get background color of the sheet tab.
         * Returns <tt>null</tt> if no sheet tab color is set.
         *
         * @return the background color of the sheet tab
         */
        /// <summary>
        /// Get or set background color of the sheet tab.
        /// The value is null if no sheet tab color is set.
        /// </summary>
        public XSSFColor TabColor
        {
            get
            {
                CT_SheetPr pr = worksheet.sheetPr;
                if (pr == null) pr = worksheet.AddNewSheetPr();
                if (!pr.IsSetTabColor())
                {
                    return null;
                }
                return new XSSFColor(pr.tabColor);
            }
            set
            {
                CT_SheetPr pr = worksheet.sheetPr;
                if (pr == null) pr = worksheet.AddNewSheetPr();
                pr.tabColor = value.GetCTColor();
            }
        }

        #region ISheet Members


        public IDrawing DrawingPatriarch
        {
            get
            {
                if (drawing == null)
                {
                    NPOI.OpenXmlFormats.Spreadsheet.CT_Drawing ctDrawing = GetCTDrawing();
                    if (ctDrawing == null)
                    {
                        return null;
                    }

                    foreach (RelationPart rp in RelationParts)
                    {
                        POIXMLDocumentPart p = rp.DocumentPart;
                        if (p is XSSFDrawing)
                        {
                            XSSFDrawing dr = (XSSFDrawing)p;
                            String drId = rp.Relationship.Id;
                            if (drId.Equals(ctDrawing.id))
                            {
                                drawing = dr;
                                break;
                            }
                            break;
                        }
                    }
                }
                return drawing;
            }
        }

        public IEnumerator GetEnumerator()
        {
            return _rows.Values.GetEnumerator();
        }

        public IEnumerator GetRowEnumerator()
        {
            return GetEnumerator();
        }

        public bool IsActive
        {
            get
            {
                return IsSelected;
            }
            set
            {
                IsSelected = value;
            }
        }

        public bool IsMergedRegion(CellRangeAddress mergedRegion)
        {
            if (worksheet.mergeCells == null || worksheet.mergeCells.mergeCell == null)
                return false;
            foreach (CT_MergeCell mc in worksheet.mergeCells.mergeCell)
            {
                if (!string.IsNullOrEmpty(mc.@ref))
                {
                    CellRangeAddress range = CellRangeAddress.ValueOf(mc.@ref);
                    if (range.FirstColumn <= mergedRegion.FirstColumn
                     && range.LastColumn >= mergedRegion.LastColumn
                     && range.FirstRow <= mergedRegion.FirstRow
                     && range.LastRow >= mergedRegion.LastRow)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        public void SetActive(bool value)
        {
            this.IsSelected = value;
        }

        public void SetActiveCellRange(List<CellRangeAddress8Bit> cellranges, int activeRange, int activeRow, int activeColumn)
        {
            throw new NotImplementedException();
        }

        public void SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            throw new NotImplementedException();
        }


        public short TabColorIndex
        {
            get
            {
                throw new NotImplementedException("Use XSSFSheet.TabColor instead");
            }
            set
            {
                throw new NotImplementedException("Use XSSFSheet.TabColor instead");
            }
        }

        public bool IsRightToLeft
        {
            get
            {
                CT_SheetView view = this.GetDefaultSheetView();
                return view == null ? false : view.rightToLeft;
            }
            set
            {
                CT_SheetView view = this.GetDefaultSheetView();
                view.rightToLeft = value;
            }
        }

        #endregion


        public IRow CopyRow(int sourceIndex, int targetIndex)
        {
            return SheetUtil.CopyRow(this, sourceIndex, targetIndex);
        }

        public CellRangeAddress RepeatingRows
        {
            get
            {
                return GetRepeatingRowsOrColums(true);
            }
            set
            {
                CellRangeAddress columnRangeRef = RepeatingColumns;
                SetRepeatingRowsAndColumns(value, columnRangeRef);
            }
        }


        public CellRangeAddress RepeatingColumns
        {
            get
            {
                return GetRepeatingRowsOrColums(false);
            }
            set
            {
                CellRangeAddress rowRangeRef = RepeatingRows;
                SetRepeatingRowsAndColumns(rowRangeRef, value);
            }
        }
        private CT_Pane Pane
        {
            get
            {
                if (GetDefaultSheetView().pane == null)
                {
                    GetDefaultSheetView().AddNewPane();
                }
                return GetDefaultSheetView().pane;
            }
        }
        public void ShowInPane(int toprow, int leftcol)
        {
            CellReference cellReference = new CellReference(toprow, leftcol);
            String cellRef = cellReference.FormatAsString();
            Pane.topLeftCell = cellRef;
        }
        private void SetRepeatingRowsAndColumns(
            CellRangeAddress rowDef, CellRangeAddress colDef)
        {
            int col1 = -1;
            int col2 = -1;
            int row1 = -1;
            int row2 = -1;

            if (rowDef != null)
            {
                row1 = rowDef.FirstRow;
                row2 = rowDef.LastRow;
                if ((row1 == -1 && row2 != -1)
                    || row1 < -1 || row2 < -1 || row1 > row2)
                {
                    throw new ArgumentException("Invalid row range specification");
                }
            }
            if (colDef != null)
            {
                col1 = colDef.FirstColumn;
                col2 = colDef.LastColumn;
                if ((col1 == -1 && col2 != -1)
                    || col1 < -1 || col2 < -1 || col1 > col2)
                {
                    throw new ArgumentException(
                        "Invalid column range specification");
                }
            }

            int sheetIndex = Workbook.GetSheetIndex(this);

            bool removeAll = rowDef == null && colDef == null;
            XSSFWorkbook xwb = Workbook as XSSFWorkbook;
            if (xwb == null)
                throw new RuntimeException("Workbook should not be null");
            XSSFName name = xwb.GetBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
            if (removeAll)
            {
                if (name != null)
                {
                    xwb.RemoveName(name);
                }
                return;
            }
            if (name == null)
            {
                name = xwb.CreateBuiltInName(
                    XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
            }

            String reference = GetReferenceBuiltInRecord(
                name.SheetName, col1, col2, row1, row2);
            name.RefersToFormula = (reference);

            // If the print setup isn't currently defined, then add it
            //  in but without printer defaults
            // If it's already there, leave it as-is!
            if (worksheet.IsSetPageSetup() && worksheet.IsSetPageMargins())
            {
                // Everything we need is already there
            }
            else
            {
                // Have initial ones put in place
                PrintSetup.ValidSettings = (false);
            }
        }

        private static String GetReferenceBuiltInRecord(
            String sheetName, int startC, int endC, int startR, int endR)
        {
            // Excel example for built-in title: 
            //   'second sheet'!$E:$F,'second sheet'!$2:$3

            CellReference colRef =
              new CellReference(sheetName, 0, startC, true, true);
            CellReference colRef2 =
              new CellReference(sheetName, 0, endC, true, true);
            CellReference rowRef =
              new CellReference(sheetName, startR, 0, true, true);
            CellReference rowRef2 =
              new CellReference(sheetName, endR, 0, true, true);

            String escapedName = SheetNameFormatter.Format(sheetName);

            String c = "";
            String r = "";

            if (startC != -1 || endC != -1)
            {
                String col1 = colRef.CellRefParts[2];
                String col2 = colRef2.CellRefParts[2];
                c = escapedName + "!$" + col1 + ":$" + col2;
            }

            if (startR != -1 || endR != -1)
            {
                String row1 = rowRef.CellRefParts[1];
                String row2 = rowRef2.CellRefParts[1];
                if (!row1.Equals("0") && !row2.Equals("0"))
                {
                    r = escapedName + "!$" + row1 + ":$" + row2;
                }
            }

            StringBuilder rng = new StringBuilder();
            rng.Append(c);
            if (rng.Length > 0 && r.Length > 0)
            {
                rng.Append(',');
            }
            rng.Append(r);
            return rng.ToString();
        }


        private CellRangeAddress GetRepeatingRowsOrColums(bool rows)
        {
            int sheetIndex = Workbook.GetSheetIndex(this);
            XSSFWorkbook xwb = Workbook as XSSFWorkbook;
            if (xwb == null)
                throw new RuntimeException("Workbook should not be null");
            XSSFName name = xwb.GetBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
            if (name == null)
            {
                return null;
            }
            String refStr = name.RefersToFormula;
            if (refStr == null)
            {
                return null;
            }
            String[] parts = refStr.Split(",".ToCharArray());
            int maxRowIndex = SpreadsheetVersion.EXCEL2007.LastRowIndex;
            int maxColIndex = SpreadsheetVersion.EXCEL2007.LastColumnIndex;
            foreach (String part in parts)
            {
                CellRangeAddress range = CellRangeAddress.ValueOf(part);
                if ((range.FirstColumn == 0
                    && range.LastColumn == maxColIndex)
                    || (range.FirstColumn == -1
                        && range.LastColumn == -1))
                {
                    if (rows)
                    {
                        return range;
                    }
                }
                else if (range.FirstRow == 0
                  && range.LastRow == maxRowIndex
                  || (range.FirstRow == -1
                      && range.LastRow == -1))
                {
                    if (!rows)
                    {
                        return range;
                    }
                }
            }
            return null;
        }

        public ISheet CopySheet(String Name)
        {
            return CopySheet(Name, true);
        }

        public ISheet CopySheet(String name, Boolean copyStyle)
        {
            String clonedName = SheetUtil.GetUniqueSheetName(this.Workbook, name);
            XSSFSheet clonedSheet = (XSSFSheet)this.Workbook.CreateSheet(clonedName);

            try
            {
                using (MemoryStream out1 = new MemoryStream())
                {
                    this.Write(out1);
                    clonedSheet.Read(new MemoryStream(out1.ToArray()));
                }
            }
            catch (IOException e)
            {
                throw new POIXMLException("Failed to clone sheet", e);
            }

            CT_Worksheet ct = clonedSheet.GetCTWorksheet();
            if (ct.IsSetLegacyDrawing())
            {
                logger.Log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
                ct.UnsetLegacyDrawing();
            }
            clonedSheet.IsSelected = false;

            // copy sheet's relations
            List<POIXMLDocumentPart> rels = this.GetRelations();
            // if the sheet being cloned has a drawing then remember it and re-create too
            XSSFDrawing dg = null;
            foreach (POIXMLDocumentPart r in rels)
            {
                // do not copy the drawing relationship, it will be re-created
                if (r is XSSFDrawing)
                {
                    dg = (XSSFDrawing)r;
                    continue;
                }
                //skip printerSettings.bin part
                if (r.GetPackagePart().PartName.Name == "/xl/printerSettings/printerSettings1.bin")
                    continue;
                PackageRelationship rel = r.GetPackageRelationship();
                clonedSheet.GetPackagePart().AddRelationship(
                    rel.TargetUri, (TargetMode)rel.TargetMode, rel.RelationshipType);
                clonedSheet.AddRelation(rel.Id, r);
            }

            // copy hyperlinks
            clonedSheet.hyperlinks = new List<XSSFHyperlink>(hyperlinks);

            // clone the sheet drawing along with its relationships
            if (dg != null)
            {
                if (ct.IsSetDrawing())
                {
                    // unset the existing reference to the drawing,
                    // so that subsequent call of clonedSheet.createDrawingPatriarch() will create a new one
                    ct.UnsetDrawing();
                }
                XSSFDrawing clonedDg = clonedSheet.CreateDrawingPatriarch() as XSSFDrawing;
                // copy drawing contents
                clonedDg.GetCTDrawing().Set(dg.GetCTDrawing());

                clonedDg = clonedSheet.CreateDrawingPatriarch() as XSSFDrawing;

                // Clone drawing relations
                List<POIXMLDocumentPart> srcRels = dg.GetRelations();
                foreach (POIXMLDocumentPart rel in srcRels)
                {
                    PackageRelationship relation = rel.GetPackageRelationship();
                    clonedDg.AddRelation(relation.Id, rel);
                    clonedDg
                            .GetPackagePart()
                            .AddRelationship(relation.TargetUri, relation.TargetMode.Value,
                                    relation.RelationshipType, relation.Id);
                }
            }
            return clonedSheet;
        }

        private void CopySheetImages(XSSFWorkbook destWorkbook, XSSFSheet destSheet)
        {
            XSSFDrawing sheetDrawing = GetDrawingPatriarch();
            if (sheetDrawing != null)
            {
                IDrawing destDraw = destSheet.CreateDrawingPatriarch();
                List<POIXMLDocumentPart> sheetPictures = sheetDrawing.GetRelations();
                Dictionary<string, uint> pictureIdMapping = new Dictionary<string, uint>();
                foreach (OpenXmlFormats.Dml.Spreadsheet.IEG_Anchor anchor in sheetDrawing.GetCTDrawing().CellAnchors)
                {
                    OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor cellAnchor = anchor as OpenXmlFormats.Dml.Spreadsheet.CT_TwoCellAnchor;
                    if (cellAnchor != null)
                    {
                        XSSFClientAnchor newAnchor = new XSSFClientAnchor((int)cellAnchor.from.colOff, (int)cellAnchor.from.rowOff,
                            (int)cellAnchor.to.colOff, (int)cellAnchor.to.rowOff, cellAnchor.from.col, cellAnchor.from.row, cellAnchor.to.col, cellAnchor.to.row);
                        if (cellAnchor.editAsSpecified)
                        {
                            switch (cellAnchor.editAs)
                            {
                                case OpenXmlFormats.Dml.Spreadsheet.ST_EditAs.twoCell:
                                    newAnchor.AnchorType = AnchorType.MoveAndResize;
                                    break;
                                case OpenXmlFormats.Dml.Spreadsheet.ST_EditAs.oneCell:
                                    newAnchor.AnchorType = AnchorType.MoveDontResize;
                                    break;
                                case OpenXmlFormats.Dml.Spreadsheet.ST_EditAs.absolute:
                                case OpenXmlFormats.Dml.Spreadsheet.ST_EditAs.NONE:
                                default:
                                    newAnchor.AnchorType = AnchorType.DontMoveAndResize;
                                    break;
                            }
                        }

                        string oldPictureId = anchor.picture.blipFill.blip.embed;
                        if (!pictureIdMapping.ContainsKey(oldPictureId))
                        {
                            XSSFPictureData srcPic = FindPicture(sheetPictures, oldPictureId);
                            if (srcPic != null && srcPic.PictureType != PictureType.None)
                            {
                                pictureIdMapping.Add(oldPictureId, (uint)destWorkbook.AddPicture(srcPic.Data, srcPic.PictureType));
                            }
                            else
                            {
                                continue; //Unable to find this picture, so skip it
                            }
                        }
                        destDraw.CreatePicture(newAnchor, (int)pictureIdMapping[oldPictureId]);
                    }
                }
            }
        }
        private XSSFPictureData FindPicture(List<POIXMLDocumentPart> sheetPictures, string id)
        {
            foreach (POIXMLDocumentPart item in sheetPictures)
            {
                if (item.GetPackageRelationship().Id == id)
                {
                    return item as XSSFPictureData;
                }
            }
            return null;
        }
        public void CopyTo(IWorkbook dest, string name, bool copyStyle, bool keepFormulas)
        {
            StylesTable styles = ((XSSFWorkbook)dest).GetStylesSource();
            if (copyStyle && Workbook.NumberOfFonts > 0)
            {
                foreach (var font in((XSSFWorkbook)Workbook).GetStylesSource().GetFonts())
                {
                    styles.PutFont(font);  //TODO::create real font mapping, the correct logic may be wrong
                }
            }
            XSSFSheet newSheet = (XSSFSheet)dest.CreateSheet(name);
            newSheet.sheet.state = sheet.state;
            IDictionary<Int32, ICellStyle> styleMap = (copyStyle) ? new Dictionary<Int32, ICellStyle>() : null;
            for (int i = FirstRowNum; i <= LastRowNum; i++)
            {
                XSSFRow srcRow = (XSSFRow)GetRow(i);
                XSSFRow destRow = (XSSFRow)newSheet.CreateRow(i);
                if (srcRow != null)
                {
                    CopyRow(this, newSheet, srcRow, destRow, styleMap, keepFormulas);
                }
            }
            List<CT_Cols> srcCols = worksheet.GetColsList();
            List<CT_Cols> dstCols = newSheet.worksheet.GetColsList();
            dstCols.Clear(); //Should already be empty since this is a new sheet.
            foreach (CT_Cols srcCol in srcCols)
            {
                CT_Cols dstCol = new CT_Cols();
                foreach (var column in srcCol.col)
                {
                    dstCol.col.Add(column.Copy());
                }
                dstCols.Add(dstCol);
            }
            newSheet.ForceFormulaRecalculation = true;
            newSheet.PrintSetup.Landscape = PrintSetup.Landscape;
            newSheet.PrintSetup.HResolution = PrintSetup.HResolution;
            newSheet.PrintSetup.VResolution = PrintSetup.VResolution;
            newSheet.SetMargin(MarginType.LeftMargin, GetMargin(MarginType.LeftMargin));
            newSheet.SetMargin(MarginType.RightMargin, GetMargin(MarginType.RightMargin));
            newSheet.SetMargin(MarginType.TopMargin, GetMargin(MarginType.TopMargin));
            newSheet.SetMargin(MarginType.BottomMargin, GetMargin(MarginType.BottomMargin));
            newSheet.PrintSetup.HeaderMargin = PrintSetup.HeaderMargin;
            newSheet.PrintSetup.FooterMargin = PrintSetup.FooterMargin;
            newSheet.Header.Left = Header.Left;
            newSheet.Header.Center = Header.Center;
            newSheet.Header.Right = Header.Right;
            newSheet.Footer.Left = Footer.Left;
            newSheet.Footer.Center = Footer.Center;
            newSheet.Footer.Right = Footer.Right;
            newSheet.PrintSetup.Scale = PrintSetup.Scale;
            newSheet.PrintSetup.FitHeight = PrintSetup.FitHeight;
            newSheet.PrintSetup.FitWidth = PrintSetup.FitWidth;
            newSheet.DisplayGridlines = DisplayGridlines;
            if (worksheet.IsSetSheetPr())
            {
                newSheet.worksheet.sheetPr = worksheet.sheetPr.Clone();
            }
            if (GetDefaultSheetView().pane != null)
            {
                var oldPane = GetDefaultSheetView().pane;
                var newPane = newSheet.GetPane();
                newPane.activePane = oldPane.activePane;
                newPane.state = oldPane.state;
                newPane.topLeftCell = oldPane.topLeftCell;
                newPane.xSplit = oldPane.xSplit;
                newPane.ySplit = oldPane.ySplit;
            }
            CopySheetImages(dest as XSSFWorkbook, newSheet);
        }
        private static void CopyRow(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow, IDictionary<Int32, ICellStyle> styleMap, bool keepFormulas)
        {
            destRow.Height = srcRow.Height;
            if (!srcRow.GetCTRow().IsSetCustomHeight())
            {
                //Copying height sets the custom height flag, but Excel will set a value for height even if it's auto-sized.
                destRow.GetCTRow().unSetCustomHeight();
            }
            destRow.Hidden = srcRow.Hidden;
            destRow.Collapsed = srcRow.Collapsed;
            destRow.OutlineLevel = srcRow.OutlineLevel;

            if (srcRow.FirstCellNum < 0)
            {
                return; //Row has no cells, this sometimes happens with hidden or blank rows
            }
            for (int j = srcRow.FirstCellNum; j <= srcRow.LastCellNum; j++)
            {
                XSSFCell oldCell = (XSSFCell)srcRow.GetCell(j);
                XSSFCell newCell = (XSSFCell)destRow.GetCell(j);
                if (srcSheet.Workbook == destSheet.Workbook)
                {
                    newCell = (XSSFCell)destRow.GetCell(j);
                }
                if (oldCell != null)
                {
                    if (newCell == null)
                    {
                        newCell = (XSSFCell)destRow.CreateCell(j);
                    }
                    XSSFSheet.CopyCell(oldCell, newCell, styleMap, keepFormulas);
                    CellRangeAddress mergedRegion = srcSheet.GetMergedRegion(new CellRangeAddress(srcRow.RowNum, srcRow.RowNum, (short)oldCell.ColumnIndex, (short)oldCell.ColumnIndex));
                    if (mergedRegion != null)
                    {
                        CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.FirstRow,
                                mergedRegion.LastRow, mergedRegion.FirstColumn, mergedRegion.LastColumn);

                        if (!destSheet.IsMergedRegion(newMergedRegion))
                        {
                            destSheet.AddMergedRegion(newMergedRegion);
                        }
                    }
                }
            }
        }
        private static void CopyCell(ICell oldCell, ICell newCell, IDictionary<Int32, ICellStyle> styleMap, Boolean keepFormulas)
        {
            if (styleMap != null)
            {
                if (oldCell.CellStyle != null)
                {
                    if (oldCell.Sheet.Workbook == newCell.Sheet.Workbook)
                    {
                        newCell.CellStyle = oldCell.CellStyle;
                    }
                    else
                    {
                        int styleHashCode = oldCell.CellStyle.GetHashCode();
                        if (styleMap.ContainsKey(styleHashCode))
                        {
                            newCell.CellStyle = styleMap[styleHashCode];
                        }
                        else
                        {
                            ICellStyle newCellStyle = (ICellStyle)newCell.Sheet.Workbook.CreateCellStyle();
                            newCellStyle.CloneStyleFrom(oldCell.CellStyle);
                            newCell.CellStyle = newCellStyle;
                            styleMap.Add(styleHashCode, newCellStyle);
                        }
                    }
                }
                else
                {
                    newCell.CellStyle = null;
                }
            }
            switch (oldCell.CellType)
            {
                case CellType.String:
                    XSSFRichTextString rts = oldCell.RichStringCellValue as XSSFRichTextString;
                    newCell.SetCellValue(rts);
                    if (rts != null)
                    {
                        for (int j = 0; j < rts.NumFormattingRuns; j++)
                        {
                            int startIndex = rts.GetIndexOfFormattingRun(j);
                            int endIndex = 0;
                            if (j + 1 == rts.NumFormattingRuns)
                            {
                                endIndex = rts.Length;
                            }
                            else
                            {
                                endIndex = rts.GetIndexOfFormattingRun(j + 1);
                            }
                            IFont fr = newCell.Sheet.Workbook.CreateFont();
                            fr.CloneStyleFrom(rts.GetFontOfFormattingRun(j));
                            newCell.RichStringCellValue.ApplyFont(startIndex, endIndex, fr);
                        }
                    }
                    break;
                case CellType.Numeric:
                    newCell.SetCellValue(oldCell.NumericCellValue);
                    break;
                case CellType.Blank:
                    newCell.SetCellType(CellType.Blank);
                    break;
                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    newCell.SetCellValue(oldCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    if (keepFormulas)
                    {
                        newCell.SetCellType(CellType.Formula);
                        newCell.CellFormula = oldCell.CellFormula;
                    }
                    else
                    {
                        try
                        {
                            newCell.SetCellType(CellType.Numeric);
                            newCell.SetCellValue(oldCell.NumericCellValue);
                        }
                        catch (Exception)
                        {
                            newCell.SetCellType(CellType.String);
                            newCell.SetCellValue(oldCell.ToString());
                        }
                    }
                    break;
                default:
                    break;
            }
        }

        public XSSFWorkbook GetWorkbook()
        {
            return (XSSFWorkbook)GetParent();
        }
        /**
         * Creates an empty XSSFPivotTable and Sets up all its relationships
         * including: pivotCacheDefInition, pivotCacheRecords
         * @return returns a pivotTable
         */
        private XSSFPivotTable CreatePivotTable()
        {
            XSSFWorkbook wb = GetWorkbook();
            List<XSSFPivotTable> pivotTables = wb.PivotTables;
            int tableId = GetWorkbook().PivotTables.Count + 1;
            //Create relationship between pivotTable and the worksheet
            XSSFPivotTable pivotTable = (XSSFPivotTable)CreateRelationship(XSSFRelation.PIVOT_TABLE,
                    XSSFFactory.GetInstance(), tableId);
            pivotTable.SetParentSheet(this);
            pivotTables.Add(pivotTable);
            XSSFWorkbook workbook = GetWorkbook();

            //Create relationship between the pivot cache defintion and the workbook
            XSSFPivotCacheDefinition pivotCacheDefinition = (XSSFPivotCacheDefinition)workbook.
                    CreateRelationship(XSSFRelation.PIVOT_CACHE_DEFINITION, XSSFFactory.GetInstance(), tableId);
            String rId = workbook.GetRelationId(pivotCacheDefinition);
            //Create relationship between pivotTable and pivotCacheDefInition without creating a new instance
            PackagePart pivotPackagePart = pivotTable.GetPackagePart();
            pivotPackagePart.AddRelationship(pivotCacheDefinition.GetPackagePart().PartName,
                    TargetMode.Internal, XSSFRelation.PIVOT_CACHE_DEFINITION.Relation);

            pivotTable.SetPivotCacheDefinition(pivotCacheDefinition);

            //Create pivotCache and Sets up it's relationship with the workbook
            pivotTable.SetPivotCache(new XSSFPivotCache(workbook.AddPivotCache(rId)));

            //Create relationship between pivotcacherecord and pivotcachedefInition
            XSSFPivotCacheRecords pivotCacheRecords = (XSSFPivotCacheRecords)pivotCacheDefinition.
                    CreateRelationship(XSSFRelation.PIVOT_CACHE_RECORDS, XSSFFactory.GetInstance(), tableId);

            //Set relationships id for pivotCacheDefInition to pivotCacheRecords
            pivotTable.GetPivotCacheDefinition().GetCTPivotCacheDefinition().id = (/*setter*/pivotCacheDefinition.GetRelationId(pivotCacheRecords));

            wb.PivotTables = (/*setter*/pivotTables);

            return pivotTable;
        }

        /**
         * Create a pivot table using the AreaReference range on sourceSheet, at the given position.
         * If the source reference contains a sheet name, it must match the sourceSheet
         * @param source location of pivot data
         * @param position A reference to the top left cell where the pivot table will start
         * @param sourceSheet The sheet containing the source data, if the source reference doesn't contain a sheet name
         * @throws IllegalArgumentException if source references a sheet different than sourceSheet
         * @return The pivot table
         */
        public XSSFPivotTable CreatePivotTable(AreaReference source, CellReference position, ISheet sourceSheet)
        {
            String sourceSheetName = source.FirstCell.SheetName;
            if (sourceSheetName != null && !sourceSheetName.Equals(sourceSheet.SheetName, StringComparison.InvariantCultureIgnoreCase))
            {
                throw new ArgumentException("The area is referenced in another sheet than the "
                        + "defined source sheet " + sourceSheet.SheetName + ".");
            }
            XSSFPivotTable.IPivotTableReferenceConfigurator refConfig = new PivotTableReferenceConfigurator1(source);
            return CreatePivotTable(position, sourceSheet, refConfig);
        }
        public class PivotTableReferenceConfigurator1 : XSSFPivotTable.IPivotTableReferenceConfigurator
        {
            AreaReference source;
            public PivotTableReferenceConfigurator1(AreaReference source)
            {
                this.source = source;
            }
            public void ConfigureReference(CT_WorksheetSource wsSource)
            {
                String[] firstCell = source.FirstCell.CellRefParts;
                String firstRow = firstCell[1];
                String firstCol = firstCell[2];
                String[] lastCell = source.LastCell.CellRefParts;
                String lastRow = lastCell[1];
                String lastCol = lastCell[2];
                String ref1 = firstCol + firstRow + ':' + lastCol + lastRow; //or just source.formatAsString()
                wsSource.@ref = ref1;
            }
        }
        /**
         * Create a pivot table using the AreaReference or named/table range on sourceSheet, at the given position.
         * If the source reference contains a sheet name, it must match the sourceSheet.
         * @param sourceRef location of pivot data - mutually exclusive with SourceName
         * @param sourceName range or table name for pivot data - mutually exclusive with SourceRef
         * @param position A reference to the top left cell where the pivot table will start
         * @param sourceSheet The sheet containing the source data, if the source reference doesn't contain a sheet name
         * @throws IllegalArgumentException if source references a sheet different than sourceSheet
         * @return The pivot table
         */
        private XSSFPivotTable CreatePivotTable(CellReference position, ISheet sourceSheet, XSSFPivotTable.IPivotTableReferenceConfigurator refConfig)
        {

            XSSFPivotTable pivotTable = CreatePivotTable();
            //Creates default Settings for the pivot table
            pivotTable.SetDefaultPivotTableDefinition();

            //Set sources and references
            pivotTable.CreateSourceReferences(position, sourceSheet, refConfig);

            //Create cachefield/s and empty SharedItems - must be after creating references
            pivotTable.GetPivotCacheDefinition().CreateCacheFields(sourceSheet);
            pivotTable.CreateDefaultDataColumns();

            return pivotTable;
        }

        /**
         * Create a pivot table using the AreaReference range, at the given position.
         * If the source reference contains a sheet name, that sheet is used, otherwise this sheet is assumed as the source sheet.
         * @param source location of pivot data
         * @param position A reference to the top left cell where the pivot table will start
         * @return The pivot table
         */
        public XSSFPivotTable CreatePivotTable(AreaReference source, CellReference position)
        {
            String sourceSheetName = source.FirstCell.SheetName;
            if (sourceSheetName != null && !sourceSheetName.Equals(this.SheetName, StringComparison.InvariantCultureIgnoreCase))
            {
                XSSFSheet sourceSheet = Workbook.GetSheet(sourceSheetName) as XSSFSheet;
                return CreatePivotTable(source, position, sourceSheet);
            }
            return CreatePivotTable(source, position, this);
        }

        /**
         * Create a pivot table using the Name range reference on sourceSheet, at the given position.
         * If the source reference contains a sheet name, it must match the sourceSheet
         * @param source location of pivot data
         * @param position A reference to the top left cell where the pivot table will start
         * @param sourceSheet The sheet containing the source data, if the source reference doesn't contain a sheet name
         * @ if source references a sheet different than sourceSheet
         * @return The pivot table
         */

        public XSSFPivotTable CreatePivotTable(IName source, CellReference position, ISheet sourceSheet)
        {
            if (source.SheetName != null && !source.SheetName.Equals(sourceSheet.SheetName))
            {
                throw new ArgumentException("The named range references another sheet than the "
                        + "defined source sheet " + sourceSheet.SheetName + ".");
            }

            return CreatePivotTable(position, sourceSheet, new PivotTableReferenceConfigurator2(source));
        }
        public class PivotTableReferenceConfigurator2 : XSSFPivotTable.IPivotTableReferenceConfigurator
        {
            IName source;
            public PivotTableReferenceConfigurator2(IName source)
            {
                this.source = source;
            }
            public void ConfigureReference(CT_WorksheetSource wsSource)
            {
                wsSource.name = (source.NameName);
            }
        }
        /**
         * Create a pivot table using the Name range, at the given position.
         * If the source reference contains a sheet name, that sheet is used, otherwise this sheet is assumed as the source sheet.
         * @param source location of pivot data
         * @param position A reference to the top left cell where the pivot table will start
         * @return The pivot table
         */

        public XSSFPivotTable CreatePivotTable(IName source, CellReference position)
        {
            return CreatePivotTable(source, position, GetWorkbook().GetSheet(source.SheetName));
        }

        /**
         * Create a pivot table using the Table, at the given position.
         * Tables are required to have a sheet reference, so no additional logic around reference sheet is needed.
         * @param source location of pivot data
         * @param position A reference to the top left cell where the pivot table will start
         * @return The pivot table
         */

        public XSSFPivotTable CreatePivotTable(ITable source, CellReference position)
        {
            return CreatePivotTable(position, GetWorkbook().GetSheet(source.SheetName), new PivotTableReferenceConfigurator3(source));
        }

        public class PivotTableReferenceConfigurator3 : XSSFPivotTable.IPivotTableReferenceConfigurator
        {
            ITable source;
            public PivotTableReferenceConfigurator3(ITable source)
            {
                this.source = source;
            }
            public void ConfigureReference(CT_WorksheetSource wsSource)
            {
                wsSource.name = (source.Name);
            }
        }
        /**
         * Returns all the pivot tables for this Sheet
         */
        public List<XSSFPivotTable> GetPivotTables()
        {
            List<XSSFPivotTable> tables = new List<XSSFPivotTable>();
            foreach (XSSFPivotTable table in GetWorkbook().PivotTables)
            {
                if (table.GetParent() == this)
                {
                    tables.Add(table);
                }
            }
            return tables;
        }

        public int GetColumnOutlineLevel(int columnIndex)
        {
            CT_Col col = columnHelper.GetColumn(columnIndex, false);
            if (col == null)
            {
                return 0;
            }
            return col.outlineLevel;
        }

        public bool IsDate1904()
        {
            throw new NotImplementedException();
        }


        /**
         * Add ignored errors (usually to suppress them in the UI of a consuming
         * application).
         *
         * @param cell Cell.
         * @param ignoredErrorTypes Types of error to ignore there.
         */
        public void AddIgnoredErrors(CellReference cell, params IgnoredErrorType[] ignoredErrorTypes)
        {
            AddIgnoredErrors(cell.FormatAsString(), ignoredErrorTypes);
        }

        /**
         * Ignore errors across a range of cells.
         * 
         * @param region Range of cells.
         * @param ignoredErrorTypes Types of error to ignore there.
         */
        public void AddIgnoredErrors(CellRangeAddress region, params IgnoredErrorType[] ignoredErrorTypes)
        {
            region.Validate(SpreadsheetVersion.EXCEL2007);
            AddIgnoredErrors(region.FormatAsString(), ignoredErrorTypes);
        }

        /**
         * Returns the errors currently being ignored and the ranges
         * where they are ignored.
         *
         * @return Map of error type to the range(s) where they are ignored.
         */
        public Dictionary<IgnoredErrorType, ISet<CellRangeAddress>> GetIgnoredErrors()
        {
            Dictionary<IgnoredErrorType, ISet<CellRangeAddress>> result = new Dictionary<IgnoredErrorType, ISet<CellRangeAddress>>();
            if (worksheet.IsSetIgnoredErrors())
            {
                foreach (CT_IgnoredError err in worksheet.ignoredErrors.ignoredError)
                {
                    foreach (IgnoredErrorType errType in GetErrorTypes(err))
                    {
                        if (!result.ContainsKey(errType))
                        {
                            result.Add(errType, new HashSet<CellRangeAddress>());
                        }
                        foreach (Object ref1 in err.sqref)
                        {
                            result[errType].Add(CellRangeAddress.ValueOf(ref1.ToString()));
                        }
                    }
                }
            }
            return result;
        }

        private void AddIgnoredErrors(String ref1, params IgnoredErrorType[] ignoredErrorTypes)
        {
            CT_IgnoredErrors ctIgnoredErrors = worksheet.IsSetIgnoredErrors() ? worksheet.ignoredErrors : worksheet.AddNewIgnoredErrors();
            CT_IgnoredError ctIgnoredError = ctIgnoredErrors.AddNewIgnoredError();
            XSSFIgnoredErrorHelper.AddIgnoredErrors(ctIgnoredError, ref1, ignoredErrorTypes);
        }

        private ISet<IgnoredErrorType> GetErrorTypes(CT_IgnoredError err)
        {
            ISet<IgnoredErrorType> result = new HashSet<IgnoredErrorType>();

            foreach (IgnoredErrorType errType in IgnoredErrorTypeValues.Values)
            {
                if (XSSFIgnoredErrorHelper.IsSet(errType, err))
                {
                    result.Add(errType);
                }
            }
            return result;
        }

    }

}
