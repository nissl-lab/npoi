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

using NPOI.HSSF.Util;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.POIFS.Crypt;
using NPOI.SS;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel.Helpers;
using SixLabors.Fonts;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text; 
using Cysharp.Text;
using System.Xml;
using CT_Shape = NPOI.OpenXmlFormats.Vml.CT_Shape;
using ST_EditAs = NPOI.OpenXmlFormats.Dml.Spreadsheet.ST_EditAs;

namespace NPOI.XSSF.UserModel
{
    /// <summary>
    /// High level representation of a SpreadsheetML worksheet. Sheets are the 
    /// central structures within a workbook, and are where a user does most of
    /// his spreadsheet work. The most common type of sheet is the worksheet, 
    /// which is represented as a grid of cells.Worksheet cells can contain 
    /// text, numbers, dates, and formulas. Cells can also be formatted.
    /// </summary>
    public partial class XSSFSheet : POIXMLDocumentPart, ISheet
    {
        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(XSSFSheet));

        private static readonly double DEFAULT_ROW_HEIGHT = 15.0;
        private static readonly double DEFAULT_MARGIN_HEADER = 0.3;
        private static readonly double DEFAULT_MARGIN_FOOTER = 0.3;
        private static readonly double DEFAULT_MARGIN_TOP = 0.75;
        private static readonly double DEFAULT_MARGIN_BOTTOM = 0.75;
        private static readonly double DEFAULT_MARGIN_LEFT = 0.7;
        private static readonly double DEFAULT_MARGIN_RIGHT = 0.7;
        public static int TWIPS_PER_POINT = 20;

        //TODO make the two variable below private!
        internal CT_Sheet sheet;
        internal CT_Worksheet worksheet;

        private readonly SortedList<int, XSSFRow> _rows = new SortedList<int, XSSFRow>();
        private readonly SortedList<int, XSSFColumn> _columns = new SortedList<int, XSSFColumn>();
        private List<XSSFHyperlink> hyperlinks;
        private ColumnHelper columnHelper;
        private CommentsTable sheetComments;

        /// <summary>
        /// cache of master shared formulas in this sheet. Master shared 
        /// formula is the first formula in a group of shared formulas is saved
        /// in the f element.
        /// </summary>
        private Dictionary<int, CT_CellFormula> sharedFormulas;

        private Dictionary<string, XSSFTable> tables;
        private List<CellRangeAddress> arrayFormulas;
        private readonly XSSFDataValidationHelper dataValidationHelper;
        private XSSFDrawing drawing = null;

        private CT_Pane Pane
        {
            get
            {
                if(GetDefaultSheetView().pane == null)
                {
                    GetDefaultSheetView().AddNewPane();
                }

                return GetDefaultSheetView().pane;
            }
        }

        #region Public properties

        /// <summary>
        /// Returns the parent XSSFWorkbook
        /// </summary>
        public IWorkbook Workbook
        {
            get
            {
                return (XSSFWorkbook) GetParent();

            }
        }

        /// <summary>
        /// Returns the name of this sheet
        /// </summary>
        public string SheetName
        {
            get
            {
                return sheet.name;
            }
        }

        /// <summary>
        /// Vertical page break information used for print layout view, page 
        /// layout view, drawing print breaksin normal view, and for printing 
        /// the worksheet.
        /// </summary>
        // YK: GetXYZArray() array accessors are deprecated in xmlbeans with
        // JDK 1.5 support
        public int[] ColumnBreaks
        {
            get
            {
                if(!worksheet.IsSetColBreaks() || worksheet.colBreaks.sizeOfBrkArray() == 0)
                {
                    return Array.Empty<int>();
                }

                List<CT_Break> brkArray = worksheet.colBreaks.brk;

                int[] breaks = new int[brkArray.Count];
                for(int i = 0; i < brkArray.Count; i++)
                {
                    CT_Break brk = brkArray[i];
                    breaks[i] = (int) brk.id - 1;
                }

                return breaks;
            }
        }

        /// <summary>
        /// Get the default column width for the sheet (if the columns do not 
        /// define their own width) in characters.
        /// </summary>
        public double DefaultColumnWidth
        {
            get
            {
                CT_SheetFormatPr pr = worksheet.sheetFormatPr;
                return (pr == null || pr.defaultColWidth == 0.0) ? 8.43 : pr.defaultColWidth;
            }
            set
            {
                var pr = GetSheetTypeSheetFormatPr();
                pr.defaultColWidth = value;
                pr.baseColWidth = 0;
            }
        }

        /// <summary>
        /// Get the default row height for the sheet (if the rows do not define 
        /// their own height) in twips(1/20 of a point)
        /// </summary>
        public short DefaultRowHeight
        {
            get
            {
                return (short) ((decimal) DefaultRowHeightInPoints * TWIPS_PER_POINT);
            }
            set
            {
                DefaultRowHeightInPoints = (float) value / TWIPS_PER_POINT;
            }
        }

        /// <summary>
        /// Get the default row height for the sheet measued in point size 
        /// (if the rows do not define their own height).
        /// </summary>
        public float DefaultRowHeightInPoints
        {
            get
            {
                CT_SheetFormatPr pr = worksheet.sheetFormatPr;
                return (float) (pr == null ? 0 : pr.defaultRowHeight);
            }
            set
            {
                CT_SheetFormatPr pr = GetSheetTypeSheetFormatPr();
                pr.defaultRowHeight = value;
                pr.customHeight = true;
            }
        }

        /// <summary>
        /// Whether the text is displayed in right-to-left mode in the window
        /// </summary>
        public bool RightToLeft
        {
            get
            {
                CT_SheetView view = GetDefaultSheetView();
                return view != null && view.rightToLeft;
            }
            set
            {
                CT_SheetView view = GetDefaultSheetView();
                view.rightToLeft = value;
            }
        }

        /// <summary>
        /// Get whether to display the guts or not, default value is true
        /// </summary>
        public bool DisplayGuts
        {
            get
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_OutlinePr outlinePr = sheetPr.outlinePr ?? new CT_OutlinePr();
                return outlinePr.showOutlineSymbols;
            }
            set
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_OutlinePr outlinePr = sheetPr.outlinePr ?? sheetPr.AddNewOutlinePr();
                outlinePr.showOutlineSymbols = value;
            }
        }

        /// <summary>
        /// Gets the flag indicating whether the window should show 0 (zero) 
        /// in cells Containing zero value. When false, cells with zero value 
        /// appear blank instead of Showing the number zero.
        /// </summary>
        public bool DisplayZeros
        {
            get
            {
                CT_SheetView view = GetDefaultSheetView();
                return view == null || view.showZeros;
            }
            set
            {
                CT_SheetView view = GetSheetTypeSheetView();
                view.showZeros = value;
            }
        }

        /// <summary>
        /// Gets the first row on the sheet
        /// </summary>
        public int FirstRowNum
        {
            get
            {
                if(_rows.Count == 0)
                {
                    return 0;
                }
                else
                {
                    foreach(int key in _rows.Keys)
                    {
                        return key;
                    }

                    throw new ArgumentOutOfRangeException();
                }
            }
        }

        /// <summary>
        /// Gets the first column on the sheet
        /// </summary>
        public int FirstColumnNum
        {
            get
            {
                if(_columns.Count == 0)
                {
                    return 0;
                }
                else
                {
                    foreach(int key in _columns.Keys)
                    {
                        return key;
                    }

                    throw new ArgumentOutOfRangeException();
                }
            }
        }

        /// <summary>
        /// Flag indicating whether the Fit to Page print option is enabled.
        /// </summary>
        public bool FitToPage
        {
            get
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_PageSetUpPr psSetup = (sheetPr == null || !sheetPr.IsSetPageSetUpPr())
                    ? new CT_PageSetUpPr()
                    : sheetPr.pageSetUpPr;
                return psSetup.fitToPage;
            }
            set
            {
                GetSheetTypePageSetUpPr().fitToPage = value;
            }
        }

        /// <summary>
        /// Returns the default footer for the sheet, creating one as needed. 
        /// You may also want to look at <see cref="FirstFooter"/>, 
        /// <see cref="OddFooter"/> and <see cref="EvenFooter"/>
        /// </summary>
        public IFooter Footer
        {
            get
            {
                // The default footer is an odd footer
                return OddFooter;
            }
        }

        /// <summary>
        /// Returns the default header for the sheet, creating one as needed.
        /// You may also want to look at <see cref="FirstFooter"/>, 
        /// <see cref="OddFooter"/> and <see cref="EvenFooter"/>
        /// </summary>
        public IHeader Header
        {
            get
            {
                // The default header is an odd header
                return OddHeader;
            }
        }

        /// <summary>
        /// Returns the odd footer. Used on all pages unless other footers
        /// also present, when used on only odd pages.
        /// </summary>
        public IFooter OddFooter
        {
            get
            {
                return new XSSFOddFooter(GetSheetTypeHeaderFooter());
            }
        }

        /// <summary>
        /// Returns the even footer. Not there by default, but when Set, 
        /// used on even pages.
        /// </summary>
        public IFooter EvenFooter
        {
            get
            {
                return new XSSFEvenFooter(GetSheetTypeHeaderFooter());
            }
        }

        /// <summary>
        /// Returns the first page footer. Not there by default, but when 
        /// Set, used on the first page.
        /// </summary>
        public IFooter FirstFooter
        {
            get
            {
                return new XSSFFirstFooter(GetSheetTypeHeaderFooter());
            }
        }

        /// <summary>
        /// Returns the odd header. Used on all pages unless other headers 
        /// also present, when used on only odd pages.
        /// </summary>
        public IHeader OddHeader
        {
            get
            {
                return new XSSFOddHeader(GetSheetTypeHeaderFooter());
            }
        }

        /// <summary>
        /// Returns the even header. Not there by default, but when 
        /// Set, used on even pages.
        /// </summary>
        public IHeader EvenHeader
        {
            get
            {
                return new XSSFEvenHeader(GetSheetTypeHeaderFooter());
            }
        }

        /// <summary>
        /// Returns the first page header. Not there by default, but when 
        /// Set, used on the first page.
        /// </summary>
        public IHeader FirstHeader
        {
            get
            {
                return new XSSFFirstHeader(GetSheetTypeHeaderFooter());
            }
        }

        /// <summary>
        /// Determine whether printed output for this sheet will be 
        /// horizontally centered.
        /// </summary>
        public bool HorizontallyCenter
        {
            get
            {
                CT_PrintOptions opts = worksheet.printOptions;
                return opts != null && opts.horizontalCentered;
            }
            set
            {
                CT_PrintOptions opts = worksheet.IsSetPrintOptions()
                    ? worksheet.printOptions
                    : worksheet.AddNewPrintOptions();
                opts.horizontalCentered = value;

            }
        }

        public int LastRowNum
        {
            get
            {
                return _rows.Count == 0 ? 0 : GetLastKey(_rows.Keys);
            }
        }

        public int LastColumnNum
        {
            get
            {
                return _columns.Count == 0 ? 0 : GetLastKey(_columns.Keys);
            }
        }

        /// <summary>
        /// Returns the list of merged regions. If you want multiple regions, 
        /// this is faster than calling {@link #getMergedRegion(int)} each time.
        /// </summary>
        public List<CellRangeAddress> MergedRegions
        {
            get
            {
                List<CellRangeAddress> addresses = new List<CellRangeAddress>();
                CT_MergeCells ctMergeCells = worksheet.mergeCells;
                if(ctMergeCells == null)
                {
                    return addresses;
                }

                foreach(CT_MergeCell ctMergeCell in ctMergeCells.mergeCell)
                {
                    string ref1 = ctMergeCell.@ref;
                    addresses.Add(CellRangeAddress.ValueOf(ref1));
                }

                return addresses;
            }
        }

        /// <summary>
        /// Returns the number of merged regions defined in this worksheet
        /// </summary>
        public int NumMergedRegions
        {
            get
            {
                CT_MergeCells ctMergeCells = worksheet.mergeCells;
                return ctMergeCells != null
                    ? ctMergeCells.sizeOfMergeCellArray()
                    : 0;
            }
        }

        public int NumHyperlinks
        {
            get
            {
                return hyperlinks.Count;
            }
        }

        /// <summary>
        /// Returns the information regarding the currently configured pane 
        /// (split or freeze).
        /// </summary>
        public PaneInformation PaneInformation
        {
            get
            {
                CT_Pane pane = GetDefaultSheetView().pane;
                // no pane configured
                if(pane == null)
                {
                    return null;
                }

                CellReference cellRef = pane.IsSetTopLeftCell() ? new CellReference(pane.topLeftCell) : null;
                return new PaneInformation((short) pane.xSplit,
                    (short) pane.ySplit,
                    cellRef == null
                        ? (short) 0
                        : (short) cellRef.Row,
                    cellRef == null
                        ? (short) 0
                        : cellRef.Col,
                    (byte) pane.activePane,
                    pane.state == ST_PaneState.frozen);
            }
        }

        /// <summary>
        /// Returns the number of phsyically defined rows (NOT the number of 
        /// rows in the sheet)
        /// </summary>
        public int PhysicalNumberOfRows
        {
            get
            {
                return _rows.Count;
            }
        }

        /// <summary>
        /// Returns the number of phsyically defined columns (NOT the number of 
        /// columns in the sheet)
        /// </summary>
        public int PhysicalNumberOfColumns
        {
            get
            {
                return _columns.Count;
            }
        }

        /// <summary>
        /// Gets the print Setup object.
        /// </summary>
        public IPrintSetup PrintSetup
        {
            get
            {
                return new XSSFPrintSetup(worksheet);
            }
        }

        /// <summary>
        /// Answer whether protection is enabled or disabled
        /// </summary>
        public bool Protect
        {
            get
            {
                return IsSheetLocked;
            }
        }

        /// <summary>
        /// Horizontal page break information used for print layout view, page 
        /// layout view, drawing print breaks in normal view, and for printing 
        /// the worksheet.
        /// </summary>
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public int[] RowBreaks
        {
            get
            {
                if(!worksheet.IsSetRowBreaks() || worksheet.rowBreaks.sizeOfBrkArray() == 0)
                {
                    return Array.Empty<int>();
                }

                List<CT_Break> brkArray = worksheet.rowBreaks.brk;
                int[] breaks = new int[brkArray.Count];
                for(int i = 0; i < brkArray.Count; i++)
                {
                    CT_Break brk = brkArray[i];
                    breaks[i] = (int) brk.id - 1;
                }

                return breaks;
            }
        }

        /// <summary>
        /// Flag indicating whether summary rows appear below detail in an 
        /// outline, when Applying an outline. When true a summary row is 
        /// inserted below the detailed data being summarized and a new outline
        /// level is established on that row. When false a summary row is 
        /// inserted above the detailed data being summarized and a new outline
        /// level is established on that row.
        /// </summary>
        public bool RowSumsBelow
        {
            get
            {
                CT_SheetPr sheetPr = worksheet.sheetPr;
                CT_OutlinePr outlinePr = (sheetPr != null && sheetPr.IsSetOutlinePr())
                    ? sheetPr.outlinePr
                    : null;
                return outlinePr == null || outlinePr.summaryBelow;
            }
            set
            {
                EnsureOutlinePr().summaryBelow = value;
            }
        }

        /// <summary>
        /// When true a summary column is inserted to the right of the detailed
        /// data being summarized and a new outline level is established on 
        /// that column. When false a summary column is inserted to the left of
        /// the detailed data being summarized and a new outline level is 
        /// established on that column.
        /// </summary>
        public bool RowSumsRight
        {
            get
            {
                CT_SheetPr sheetPr = worksheet.sheetPr;
                CT_OutlinePr outlinePr = (sheetPr != null && sheetPr.IsSetOutlinePr())
                    ? sheetPr.outlinePr
                    : new CT_OutlinePr();
                return outlinePr.summaryRight;
            }
            set
            {
                EnsureOutlinePr().summaryRight = value;
            }
        }

        /// <summary>
        /// A flag indicating whether scenarios are locked when the sheet 
        /// is protected.
        /// </summary>
        public bool ScenarioProtect
        {
            get
            {
                return worksheet.IsSetSheetProtection()
                       && worksheet.sheetProtection.scenarios;
            }
        }

        public short LeftCol
        {
            get
            {
                string cellRef = GetPane().topLeftCell;
                if(cellRef == null)
                {
                    return 0;
                }

                CellReference cellReference = new CellReference(cellRef);
                return cellReference.Col;
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        /// <summary>
        /// The top row in the visible view when the sheet is first viewed 
        /// after opening it in a viewer
        /// </summary>
        public short TopRow
        {
            get
            {
                string cellRef = GetSheetTypeSheetView().topLeftCell;
                if(cellRef == null)
                {
                    return 0;
                }

                CellReference cellReference = new CellReference(cellRef);
                return (short) cellReference.Row;
            }
            set
            {
                GetSheetTypeSheetView().topLeftCell = "A" + value.ToString();
            }
        }

        /// <summary>
        /// Determine whether printed output for this sheet will be vertically 
        /// centered.
        /// </summary>
        public bool VerticallyCenter
        {
            get
            {
                CT_PrintOptions opts = worksheet.printOptions;
                return opts != null && opts.verticalCentered;
            }
            set
            {
                CT_PrintOptions opts = worksheet.IsSetPrintOptions()
                    ? worksheet.printOptions
                    : worksheet.AddNewPrintOptions();
                opts.verticalCentered = value;

            }
        }

        /// <summary>
        /// Gets the flag indicating whether this sheet should display formulas.
        /// </summary>
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

        /// <summary>
        /// Gets the flag indicating whether this sheet displays the lines 
        /// between rows and columns to make editing and Reading easier.
        /// </summary>
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

        /// <summary>
        /// Gets the flag indicating whether this sheet should display row and 
        /// column headings. Row heading are the row numbers to the side of the
        /// sheet Column heading are the letters or numbers that appear above 
        /// the columns of the sheet
        /// </summary>
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

        /// <summary>
        /// Returns whether gridlines are printed.
        /// </summary>
        public bool IsPrintGridlines
        {
            get
            {
                CT_PrintOptions opts = worksheet.printOptions;
                return opts != null && opts.gridLines;
            }
            set
            {
                CT_PrintOptions opts = worksheet.IsSetPrintOptions()
                    ? worksheet.printOptions
                    : worksheet.AddNewPrintOptions();
                opts.gridLines = value;
            }
        }

        /// <summary>
        /// Returns whether row and column headings are printed.
        /// </summary>
        public bool IsPrintRowAndColumnHeadings
        {
            get
            {
                CT_PrintOptions opts = worksheet.printOptions;
                return opts != null && opts.headings;
            }
            set
            {
                CT_PrintOptions opts = worksheet.IsSetPrintOptions()
                    ? worksheet.printOptions
                    : worksheet.AddNewPrintOptions();
                opts.headings = value;
            }
        }

        /// <summary>
        /// Whether Excel will be asked to recalculate all formulas when the 
        /// workbook is opened.
        /// </summary>
        public bool ForceFormulaRecalculation
        {
            get
            {
                if(worksheet.IsSetSheetCalcPr())
                {
                    CT_SheetCalcPr calc = worksheet.sheetCalcPr;
                    return calc.fullCalcOnLoad;
                }

                return false;
            }
            set
            {
                CT_CalcPr calcPr = (Workbook as XSSFWorkbook).GetCTWorkbook().calcPr;
                if(worksheet.IsSetSheetCalcPr())
                {
                    // Change the current Setting
                    CT_SheetCalcPr calc = worksheet.sheetCalcPr;
                    calc.fullCalcOnLoad = value;
                }
                else if(value)
                {
                    // Add the Calc block and set it
                    CT_SheetCalcPr calc = worksheet.AddNewSheetCalcPr();
                    calc.fullCalcOnLoad = value;
                }

                if(value && calcPr != null
                         && calcPr.calcMode == ST_CalcMode.manual)
                {
                    calcPr.calcMode = ST_CalcMode.auto;
                }
            }
        }

        /// <summary>
        /// Flag indicating whether the sheet displays Automatic Page Breaks.
        /// </summary>
        public bool Autobreaks
        {
            get
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_PageSetUpPr psSetup = (sheetPr == null || !sheetPr.IsSetPageSetUpPr())
                    ? new CT_PageSetUpPr()
                    : sheetPr.pageSetUpPr;
                return psSetup.autoPageBreaks;
            }
            set
            {
                CT_SheetPr sheetPr = GetSheetTypeSheetPr();
                CT_PageSetUpPr psSetup = sheetPr.IsSetPageSetUpPr()
                    ? sheetPr.pageSetUpPr
                    : sheetPr.AddNewPageSetUpPr();
                psSetup.autoPageBreaks = value;
            }
        }

        /// <summary>
        /// Returns a flag indicating whether this sheet is selected.
        /// <para>
        /// When only 1 sheet is selected and active, this value should be in 
        /// synch with the activeTab value. In case of a conflict, the Start 
        /// Part Setting wins and Sets the active sheet tab.
        /// </para>
        /// Note: multiple sheets can be selected, but only one sheet can be 
        /// active at one time.
        /// </summary>
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
                foreach(CT_SheetView view in views.sheetView)
                {
                    view.tabSelected = value;
                }
            }
        }

        /// <summary>
        /// Return location of the active cell, e.g. <code>A1</code>.
        /// </summary>
        public CellAddress ActiveCell
        {
            get
            {
                string address = GetSheetTypeSelection().activeCell;
                if(address == null)
                {
                    return null;
                }

                return new CellAddress(address);
            }
            set
            {
                string ref1 = value.FormatAsString();
                CT_Selection ctsel = GetSheetTypeSelection();
                ctsel.activeCell = ref1;
                ctsel.SetSqref(new string[] { ref1 });
            }
        }

        /// <summary>
        /// Does this sheet have any comments on it? We need to know, so we can
        /// decide about writing it to disk or not
        /// </summary>
        public bool HasComments
        {
            get
            {
                if(sheetComments == null)
                {
                    return false;
                }

                return sheetComments.GetNumberOfComments() > 0;
            }
        }

        internal int NumberOfComments
        {
            get
            {
                if(sheetComments == null)
                {
                    return 0;
                }

                return sheetComments.GetNumberOfComments();
            }
        }

        /// <summary>
        /// true when Autofilters are locked and the sheet is protected.
        /// </summary>
        public bool IsAutoFilterLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().autoFilter;
            }
        }

        /// <summary>
        /// true when Deleting columns is locked and the sheet is protected.
        /// </summary>
        public bool IsDeleteColumnsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().deleteColumns;
            }
        }

        /// <summary>
        /// true when Deleting rows is locked and the sheet is protected.
        /// </summary>
        public bool IsDeleteRowsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().deleteRows;
            }
        }

        /// <summary>
        /// true when Formatting cells is locked and the sheet is protected.
        /// </summary>
        public bool IsFormatCellsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().formatCells;
            }
        }

        /// <summary>
        /// true when Formatting columns is locked and the sheet is protected.
        /// </summary>
        public bool IsFormatColumnsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().formatColumns;
            }
        }

        /// <summary>
        /// true when Formatting rows is locked and the sheet is protected.
        /// </summary>
        public bool IsFormatRowsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().formatRows;
            }
        }

        /// <summary>
        /// true when Inserting columns is locked and the sheet is protected.
        /// </summary>
        public bool IsInsertColumnsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().insertColumns;
            }
        }

        /// <summary>
        /// true when Inserting hyperlinks is locked and the sheet is protected.
        /// </summary>
        public bool IsInsertHyperlinksLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().insertHyperlinks;
            }
        }

        /// <summary>
        /// true when Inserting rows is locked and the sheet is protected.
        /// </summary>
        public bool IsInsertRowsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().insertRows;
            }
        }

        /// <summary>
        /// true when Pivot tables are locked and the sheet is protected.
        /// </summary>
        public bool IsPivotTablesLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().pivotTables;
            }
        }

        /// <summary>
        /// true when Sorting is locked and the sheet is protected.
        /// </summary>
        public bool IsSortLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().sort;
            }
        }

        /// <summary>
        /// true when Objects are locked and the sheet is protected.
        /// </summary>
        public bool IsObjectsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().objects;
            }
        }

        /// <summary>
        /// true when Scenarios are locked and the sheet is protected.
        /// </summary>
        public bool IsScenariosLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().scenarios;
            }
        }

        /// <summary>
        /// true when Selection of locked cells is locked and the sheet is 
        /// protected.
        /// </summary>
        public bool IsSelectLockedCellsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().selectLockedCells;
            }
        }

        /// <summary>
        /// true when Selection of unlocked cells is locked and the sheet is 
        /// protected.
        /// </summary>
        public bool IsSelectUnlockedCellsLocked
        {
            get
            {
                return IsSheetLocked && SafeGetProtectionField().selectUnlockedCells;
            }
        }

        /// <summary>
        /// true when Sheet is Protected.
        /// </summary>
        public bool IsSheetLocked
        {
            get
            {
                return worksheet.IsSetSheetProtection() && SafeGetProtectionField().sheet;
            }
        }

        public ISheetConditionalFormatting SheetConditionalFormatting
        {
            get
            {
                return new XSSFSheetConditionalFormatting(this);
            }
        }

        /// <summary>
        /// Get or set background color of the sheet tab. The value is null 
        /// if no sheet tab color is set.
        /// </summary>
        public XSSFColor TabColor
        {
            get
            {
                CT_SheetPr pr = worksheet.sheetPr;
                if(pr == null)
                {
                    pr = worksheet.AddNewSheetPr();
                }

                if(!pr.IsSetTabColor())
                {
                    return null;
                }

                return new XSSFColor(pr.tabColor);
            }
            set
            {
                CT_SheetPr pr = worksheet.sheetPr;
                if(pr == null)
                {
                    pr = worksheet.AddNewSheetPr();
                }

                pr.tabColor = value.GetCTColor();
            }
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

        #endregion

        #region Constructors

        /// <summary>
        /// Creates new XSSFSheet   - called by XSSFWorkbook to create a sheet 
        /// from scratch. See <see cref="XSSFWorkbook.CreateSheet"/>
        /// </summary>
        public XSSFSheet()
            : base()
        {

            dataValidationHelper = new XSSFDataValidationHelper(this);
            OnDocumentCreate();
        }

        /// <summary>
        /// Creates an XSSFSheet representing the given namespace part and 
        /// relationship. Should only be called by XSSFWorkbook when Reading in
        /// an exisiting file.
        /// </summary>
        /// <param name="part">The namespace part that holds xml data 
        /// represenring this sheet.</param>
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

        #endregion

        #region Internal methods

        /// <summary>
        /// Initialize worksheet data when Reading in an exisiting file.
        /// </summary>
        /// <exception cref="POIXMLException"></exception>
        internal override void OnDocumentRead()
        {
            try
            {
                Read(GetPackagePart().GetInputStream());
            }
            catch(IOException e)
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
            catch(XmlException e)
            {
                throw new POIXMLException(e);
            }

            InitRows(worksheet);
            InitColumns(worksheet);

            // Look for bits we're interested in
            foreach(RelationPart rp in RelationParts)
            {
                POIXMLDocumentPart p = rp.DocumentPart;
                if(p is CommentsTable commentsTable)
                {
                    sheetComments = commentsTable;
                    //break;
                }

                if(p is XSSFTable xssfTable)
                {
                    tables[rp.Relationship.Id] = xssfTable;
                }

                if(p is XSSFPivotTable pivotTable)
                {
                    GetWorkbook().PivotTables.Add(pivotTable);
                }
            }

            // Process external hyperlinks for the sheet, if there are any
            InitHyperlinks();
        }

        /// <summary>
        /// Initialize worksheet data when creating a new sheet.
        /// </summary>
        internal override void OnDocumentCreate()
        {
            worksheet = NewSheet();
            InitRows(worksheet);
            InitColumns(worksheet);
            columnHelper = new ColumnHelper(worksheet);
            hyperlinks = new List<XSSFHyperlink>();
        }

        /// <summary>
        /// Get VML drawing for this sheet (aka 'legacy' drawig)
        /// </summary>
        /// <param name="autoCreate">if true, then a new VML drawing part 
        /// is Created</param>
        /// <returns>the VML drawing of null if the drawing was not found and 
        /// autoCreate=false</returns>
        internal XSSFVMLDrawing GetVMLDrawing(bool autoCreate)
        {
            XSSFVMLDrawing drawing = null;
            OpenXmlFormats.Spreadsheet.CT_LegacyDrawing ctDrawing = GetCTLegacyDrawing();
            if(ctDrawing == null)
            {
                if(autoCreate)
                {
                    //drawingNumber = #drawings.Count + 1
                    int drawingNumber = GetPackagePart()
                        .Package.GetPartsByContentType(XSSFRelation.VML_DRAWINGS.ContentType).Count + 1;
                    RelationPart rp = CreateRelationship(
                        XSSFRelation.VML_DRAWINGS,
                        XSSFFactory.GetInstance(),
                        drawingNumber,
                        false);
                    drawing = rp.DocumentPart as XSSFVMLDrawing;
                    string relId = rp.Relationship.Id;

                    // add CT_LegacyDrawing element which indicates that this
                    // sheet Contains drawing components built on the drawingML
                    // platform. The relationship Id references the part
                    // Containing the drawing defInitions.
                    ctDrawing = worksheet.AddNewLegacyDrawing();
                    ctDrawing.id = relId;
                }
            }
            else
            {
                //search the referenced drawing in the list of the sheet's relations
                string id = ctDrawing.id;
                foreach(RelationPart rp in RelationParts)
                {
                    POIXMLDocumentPart p = rp.DocumentPart;
                    if(p is XSSFVMLDrawing dr)
                    {
                        string drId = rp.Relationship.Id;
                        if(drId.Equals(id))
                        {
                            drawing = dr;
                            break;
                        }
                        // do not break here since drawing has not been found yet (see bug 52425)
                    }
                }

                if(drawing == null)
                {
                    logger.Log(POILogger.ERROR, "Can't find VML drawing with " +
                                                "id=" + id + " in the list of the sheet's relationships");
                }
            }

            return drawing;
        }

        /// <summary>
        /// Returns the sheet's comments object if there is one, or null if not
        /// </summary>
        /// <param name="create">create a new comments table if it does not 
        /// exist</param>
        /// <returns></returns>
        protected internal CommentsTable GetCommentsTable(bool create)
        {
            if(sheetComments == null && create)
            {
                // Try to create a comments table with the same number as the
                // sheet has (i.e. sheet 1 -> comments 1)
                try
                {
                    sheetComments = (CommentsTable) CreateRelationship(
                        XSSFRelation.SHEET_COMMENTS, XSSFFactory.GetInstance(), (int) sheet.sheetId);
                }
                catch(PartAlreadyExistsException)
                {
                    // Technically a sheet doesn't need the same number as it's
                    // comments, and clearly someone has already pinched our
                    // number! Go for the next available one instead
                    sheetComments = (CommentsTable) CreateRelationship(
                        XSSFRelation.SHEET_COMMENTS, XSSFFactory.GetInstance(), -1);
                }
            }

            return sheetComments;
        }

        /// <summary>
        /// Return a master shared formula by index
        /// </summary>
        /// <param name="sid">shared group index</param>
        /// <returns>a CT_CellFormula bean holding shared formula or 
        /// <code>null</code> if not found</returns>
        internal CT_CellFormula GetSharedFormula(int sid)
        {
            return sharedFormulas[sid];
        }

        internal void OnReadCell(XSSFCell cell)
        {
            //collect cells holding shared formulas
            CT_Cell ct = cell.GetCTCell();
            CT_CellFormula f = ct.f;
            if(f != null && f.t == ST_CellFormulaType.shared
                         && f.isSetRef() && f.Value != null)
            {
                // save a detached  copy to avoid XmlValueDisconnectedException,
                // this may happen when the master cell of a shared formula
                // is Changed
                CT_CellFormula sf = f.Copy();
                CellRangeAddress sfRef = CellRangeAddress.ValueOf(sf.@ref);
                CellReference cellRef = new CellReference(cell);
                // If the shared formula range preceeds the master cell then
                // the preceding  part is discarded, e.g. if the cell is E60
                // and the shared formula range is C60:M85 then the effective
                // range is E60:M85 see more details in
                // https://issues.apache.org/bugzilla/show_bug.cgi?id=51710
                if(cellRef.Col > sfRef.FirstColumn || cellRef.Row > sfRef.FirstRow)
                {
                    string effectiveRef = new CellRangeAddress(
                            Math.Max(cellRef.Row, sfRef.FirstRow), sfRef.LastRow,
                            Math.Max(cellRef.Col, sfRef.FirstColumn), sfRef.LastColumn)
                        .FormatAsString();
                    sf.@ref = effectiveRef;
                }

                sharedFormulas[(int) f.si] = sf;
            }

            if(f != null && f.t == ST_CellFormulaType.array && f.@ref != null)
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

        protected virtual OpenXmlFormats.Spreadsheet.CT_Drawing GetCTDrawing()
        {
            return worksheet.drawing;
        }

        protected virtual OpenXmlFormats.Spreadsheet.CT_LegacyDrawing GetCTLegacyDrawing()
        {
            return worksheet.legacyDrawing;
        }

        internal virtual void Write(Stream stream, bool leaveOpen = false)
        {
            bool setToNull = false;
            if(worksheet.sizeOfColsArray() == 1)
            {
                CT_Cols col = worksheet.GetColsArray(0);
                if(col.sizeOfColArray() == 0)
                {
                    setToNull = true;
                    // this is necessary so that we do not write an empty
                    // <cols/> item into the sheet-xml in the xlsx-file Excel
                    // complains about a corrupted file if this shows up there!
                    worksheet.SetColsArray(null);
                }
                else
                {
                    SetColWidthAttribute(col);
                }
            }

            // Now re-generate our CT_Hyperlinks, if needed
            if(hyperlinks.Count > 0)
            {
                if(worksheet.hyperlinks == null)
                {
                    worksheet.AddNewHyperlinks();
                }

                CT_Hyperlink[] ctHls
                    = new CT_Hyperlink[hyperlinks.Count];
                for(int i = 0; i < ctHls.Length; i++)
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
            else
            {
                if(worksheet.hyperlinks != null)
                {
                    int count = worksheet.hyperlinks.SizeOfHyperlinkArray();
                    for(int i = count - 1; i >= 0; i--)
                    {
                        worksheet.hyperlinks.RemoveHyperlink(i);
                    }

                    // For some reason, we have to remove the hyperlinks one by one from the CTHyperlinks array
                    // before unsetting the hyperlink array.
                    // Resetting the hyperlink array seems to break some XML nodes.
                    //worksheet.getHyperlinks().setHyperlinkArray(Array.Empty<CTHyperlink>());
                    worksheet.UnsetHyperlinks();
                }
                else
                {
                    // nothing to do
                }
            }

            foreach(XSSFRow row in _rows.Values)
            {
                row.OnDocumentWrite();
            }

            int minCell = int.MaxValue, maxCell = int.MinValue;
            foreach(XSSFRow row in _rows.Values)
            {
                // first perform the normal write actions for the row
                row.OnDocumentWrite();

                // then calculate min/max cell-numbers for the worksheet-dimension
                if(row.FirstCellNum != -1)
                {
                    minCell = Math.Min(minCell, row.FirstCellNum);
                }

                if(row.LastCellNum != -1)
                {
                    maxCell = Math.Max(maxCell, row.LastCellNum);
                }
            }

            foreach(XSSFColumn column in _columns.Values)
            {
                column.OnDocumentWrite();
            }

            // finally, if we had at least one cell we can populate the optional dimension-field
            if(minCell != int.MaxValue)
            {
                string ref1 = new CellRangeAddress(FirstRowNum, LastRowNum, minCell, maxCell).FormatAsString();
                if(worksheet.IsSetDimension())
                {
                    worksheet.dimension.@ref = ref1;
                }
                else
                {
                    worksheet.AddNewDimension().@ref = (ref1);
                }
            }

            new WorksheetDocument(worksheet).Save(stream, leaveOpen);

            // Bug 52233: Ensure that we have a col-array even if write() removed it
            if(setToNull)
            {
                worksheet.AddNewCols();
            }
        }

        internal bool IsCellInArrayFormulaContext(ICell cell)
        {
            foreach(CellRangeAddress range in arrayFormulas)
            {
                if(range.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return true;
                }
            }

            return false;
        }

        internal XSSFCell GetFirstCellInArrayFormula(ICell cell)
        {
            foreach(CellRangeAddress range in arrayFormulas)
            {
                if(range.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return (XSSFCell) GetRow(range.FirstRow)
                        .GetCell(range.FirstColumn);
                }
            }

            return null;
        }

        /// <summary>
        /// when a cell with a 'master' shared formula is removed,  the next
        /// cell in the range becomes the master
        /// </summary>
        /// <param name="cell">The cell that is removed</param>
        /// <param name="evalWb">in use, if one exists</param>
        internal void OnDeleteFormula(XSSFCell cell, XSSFEvaluationWorkbook evalWb)
        {

            CT_CellFormula f = cell.GetCTCell().f;
            if(f != null
               && f.t == ST_CellFormulaType.shared
               && f.isSetRef()
               && f.Value != null)
            {
                bool breakit = false;
                CellRangeAddress ref1 = CellRangeAddress.ValueOf(f.@ref);
                if(ref1.NumberOfCells > 1)
                {
                    for(int i = cell.RowIndex; i <= ref1.LastRow; i++)
                    {
                        XSSFRow row = (XSSFRow) GetRow(i);
                        if(row != null)
                        {
                            for(int j = cell.ColumnIndex; j <= ref1.LastColumn; j++)
                            {
                                XSSFCell nextCell = (XSSFCell) row.GetCell(j);
                                if(nextCell != null
                                   && nextCell != cell
                                   && nextCell.CellType == CellType.Formula)
                                {
                                    CT_CellFormula nextF = nextCell.GetCTCell().f;
                                    if(nextF.t == ST_CellFormulaType.shared
                                       && nextF.si == f.si)
                                    {
                                        nextF.Value = nextCell.GetCellFormula(evalWb);
                                        CellRangeAddress nextRef = new CellRangeAddress(
                                            nextCell.RowIndex, ref1.LastRow,
                                            nextCell.ColumnIndex, ref1.LastColumn);
                                        nextF.@ref = nextRef.FormatAsString();

                                        sharedFormulas[(int) nextF.si] = nextF;
                                        breakit = true;
                                        break;
                                    }
                                }
                            }

                            if(breakit)
                            {
                                break;
                            }
                        }
                    }
                }
            }
        }

        IEnumerator<IRow> IEnumerable<IRow>.GetEnumerator()
        {
            return _rows.Values.GetEnumerator();
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Provide access to the CT_Worksheet bean holding this sheet's data
        /// </summary>
        /// <returns>the CT_Worksheet bean holding this sheet's data</returns>
        public CT_Worksheet GetCTWorksheet()
        {
            return worksheet;
        }

        public ColumnHelper GetColumnHelper()
        {
            columnHelper = columnHelper ?? new ColumnHelper(worksheet);
            return columnHelper;
        }

        /// <summary>
        /// Adds a merged region of cells on a sheet.
        /// </summary>
        /// <param name="region">region to merge</param>
        /// <returns>index of this region</returns>
        /// <exception cref="ArgumentException">if region contains fewer 
        /// than 2 cells</exception>
        /// <exception cref="InvalidOperationException">if region intersects 
        /// with an existing merged region or multi-cell array formula on 
        /// this sheet</exception>
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
        /// <exception cref="ArgumentException">if region contains fewer 
        /// than 2 cells</exception>
        public int AddMergedRegionUnsafe(CellRangeAddress region)
        {
            return AddMergedRegion(region, false);
        }

        /// <summary>
        /// Verify that merged regions do not intersect multi-cell array 
        /// formulas and no merged regions intersect another merged region 
        /// in this sheet.
        /// </summary>
        public void ValidateMergedRegions()
        {
            CheckForMergedRegionsIntersectingArrayFormulas();
            CheckForIntersectingMergedRegions();
        }

        /// <summary>
        /// Adjusts the column width to fit the contents.
        /// This process can be relatively slow on large sheets, so this should
        /// normally only be called once per column, at the end of your
        /// Processing.
        /// </summary>
        /// <param name="column">the column index</param>
        public void AutoSizeColumn(int column)
        {
            AutoSizeColumn(column, false);
        }

        /// <summary>
        /// Adjusts the column width to fit the contents.
        /// This process can be relatively slow on large sheets, so this should
        /// normally only be called once per column, at the end of your
        /// Processing.
        /// </summary>
        /// <param name="column">the column index</param>
        /// <param name="useMergedCells">whether to use the contents of merged cells 
        /// when calculating the width of the column</param>
        public void AutoSizeColumn(int column, bool useMergedCells)
        {
            double width = SheetUtil.GetColumnWidth(this, column, useMergedCells);

            if(width != -1)
            {
                width *= 256;
                // The maximum column width for an individual cell is 255 characters
                int maxColumnWidth = 255 * 256;
                if(width > maxColumnWidth)
                {
                    width = maxColumnWidth;
                }

                IColumn col = GetColumn(column, true);
                col.Width = width / 256;
                col.IsBestFit = true;
            }
        }

        /// <summary>
        /// Adjusts the row height to fit the contents.
        /// This process can be relatively slow on large sheets, so this should
        /// normally only be called once per row, at the end of your
        /// Processing.
        /// </summary>
        /// <param name="row">the row index</param>
        public void AutoSizeRow(int row)
        {
            AutoSizeRow(row, false);
        }

        /// <summary>
        /// Adjusts the row height to fit the contents. This process can be 
        /// relatively slow on large sheets, so this should normally only be 
        /// called once per row, at the end of your Processing. You can specify
        /// whether the content of merged cells should be considered or 
        /// ignored. Default is to ignore merged cells.
        /// </summary>
        /// <param name="row">the row index</param>
        /// <param name="useMergedCells">whether to use the contents of merged 
        /// cells when  calculating the height of the row</param>
        public void AutoSizeRow(int row, bool useMergedCells)
        {
            IRow targetRow = GetRow(row) ?? CreateRow(row);

            double height = SheetUtil.GetRowHeight(this, row, useMergedCells);

            if(height != -1 && height != 0)
            {
                height *= 20;
                // The maximum row height for an individual cell is 409 points
                int maxRowHeight = 409 * 20;

                if(height > maxRowHeight)
                {
                    height = maxRowHeight;
                }

                targetRow.Height = (short) height;
            }
        }

        /// <summary>
        /// Return the sheet's existing Drawing, or null if there isn't yet one.
        /// Use <see cref="CreateDrawingPatriarch"/> to Get or create
        /// </summary>
        /// <returns>a SpreadsheetML Drawing</returns>
        public XSSFDrawing GetDrawingPatriarch()
        {
            OpenXmlFormats.Spreadsheet.CT_Drawing ctDrawing = GetCTDrawing();
            if(ctDrawing != null)
            {
                // Search the referenced Drawing in the list of the sheet's relations
                foreach(RelationPart rp in RelationParts)
                {
                    POIXMLDocumentPart p = rp.DocumentPart;
                    if(p is XSSFDrawing dr)
                    {
                        string drId = rp.Relationship.Id;
                        if(drId.Equals(ctDrawing.id))
                        {
                            return dr;
                        }

                        break;
                    }
                }

                logger.Log(POILogger.ERROR, "Can't find Drawing with id=" +
                                            ctDrawing.id + " in the list of the sheet's relationships");
            }

            return null;
        }

        /// <summary>
        /// Create a new SpreadsheetML Drawing. If this sheet already 
        /// Contains a Drawing - return that.
        /// </summary>
        /// <returns>a SpreadsheetML Drawing</returns>
        public IDrawing CreateDrawingPatriarch()
        {
            OpenXmlFormats.Spreadsheet.CT_Drawing ctDrawing = GetCTDrawing();
            if(ctDrawing != null)
            {
                return GetDrawingPatriarch();
            }

            // Default drawingNumber = #drawings.Count + 1
            int drawingNumber = GetPackagePart()
                .Package.GetPartsByContentType(XSSFRelation.DRAWINGS.ContentType).Count + 1;
            drawingNumber = GetNextPartNumber(XSSFRelation.DRAWINGS, drawingNumber);
            RelationPart rp = CreateRelationship(
                XSSFRelation.DRAWINGS,
                XSSFFactory.GetInstance(),
                drawingNumber,
                false);
            XSSFDrawing drawing = rp.DocumentPart as XSSFDrawing;
            string relId = rp.Relationship.Id;

            // add CT_Drawing element which indicates that this sheet Contains
            // Drawing components built on the DrawingML platform. The
            // relationship Id references the part Containing the
            // DrawingML defInitions.
            ctDrawing = worksheet.AddNewDrawing();
            ctDrawing.id = /*setter*/relId;

            // Return the newly Created Drawing
            return drawing;
        }

        /// <summary>
        /// Creates a split (freezepane). Any existing freezepane or split 
        /// pane is overwritten.
        /// </summary>
        /// <param name="colSplit">Horizonatal position of split.</param>
        /// <param name="rowSplit">Vertical position of split.</param>
        public void CreateFreezePane(int colSplit, int rowSplit)
        {
            CreateFreezePane(colSplit, rowSplit, colSplit, rowSplit);
        }

        /// <summary>
        /// Creates a split (freezepane). Any existing freezepane or split pane
        /// is overwritten. If both colSplit and rowSplit are zero then the 
        /// existing freeze pane is Removed
        /// </summary>
        /// <param name="colSplit">Horizonatal position of split.</param>
        /// <param name="rowSplit">Vertical position of split.</param>
        /// <param name="leftmostColumn">Left column visible in right pane.</param>
        /// <param name="topRow">Top row visible in bottom pane</param>
        public void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
        {
            CT_SheetView ctView = GetDefaultSheetView();

            // If both colSplit and rowSplit are zero then the existing freeze pane is Removed
            if(colSplit == 0 && rowSplit == 0)
            {

                if(ctView.IsSetPane())
                {
                    ctView.UnsetPane();
                }

                ctView.SetSelectionArray(null);
                return;
            }

            if(!ctView.IsSetPane())
            {
                ctView.AddNewPane();
            }

            CT_Pane pane = ctView.pane;

            if(colSplit > 0)
            {
                pane.xSplit = colSplit;
            }
            else
            {

                if(pane.IsSetXSplit())
                {
                    pane.UnsetXSplit();
                }
            }

            if(rowSplit > 0)
            {
                pane.ySplit = rowSplit;
            }
            else
            {
                if(pane.IsSetYSplit())
                {
                    pane.UnsetYSplit();
                }
            }

            pane.state = ST_PaneState.frozen;
            if(rowSplit == 0)
            {
                pane.topLeftCell = new CellReference(0, leftmostColumn).FormatAsString();
                pane.activePane = ST_Pane.topRight;
            }
            else if(colSplit == 0)
            {
                pane.topLeftCell = new CellReference(topRow, 0).FormatAsString();
                pane.activePane = ST_Pane.bottomLeft;
            }
            else
            {
                pane.topLeftCell = new CellReference(topRow, leftmostColumn).FormatAsString();
                pane.activePane = ST_Pane.bottomRight;
            }

            ctView.selection = null;
            CT_Selection sel = ctView.AddNewSelection();
            sel.pane = pane.activePane;
        }

        /// <summary>
        /// Create a new row within the sheet and return the high level 
        /// representation. See <see cref="RemoveRow"/>
        /// </summary>
        /// <param name="rownum">row number</param>
        /// <returns>High level <see cref="XSSFRow"/> object representing a 
        /// row in the sheet</returns>
        public virtual IRow CreateRow(int rownum)
        {
            CT_Row ctRow;
            XSSFRow prev = _rows.TryGetValue(rownum, out XSSFRow row) ? row : null;
            if (prev != null)
            {
                // the Cells in an existing row are invalidated on-purpose, in
                // order to clean up correctly, we need to call the remove, so
                // things like ArrayFormulas and CalculationChain updates are
                // done correctly. We remove the cell this way as the internal
                // cell-list is changed by the remove call and thus would cause
                // ConcurrentModificationException otherwise
                while(prev.FirstCellNum != -1)
                {
                    prev.RemoveCell(prev.GetCell(prev.FirstCellNum));
                }

                ctRow = prev.GetCTRow();
                ctRow.Set(new CT_Row());
            }
            else
            {
                if(_rows.Count == 0 || rownum > GetLastKey(_rows.Keys))
                {
                    // we can append the new row at the end
                    ctRow = worksheet.sheetData.AddNewRow();
                }
                else
                {
                    // get number of rows where row index < rownum
                    // --> this tells us where our row should go
                    int idx = HeadMapCount(_rows.Keys, rownum);
                    ctRow = worksheet.sheetData.InsertNewRow(idx);
                }
            }

            XSSFRow r = new XSSFRow(ctRow, this) { RowNum = rownum };
            _rows[rownum] = r;
            return r;
        }

        /// <summary>
        /// Create a new column within the sheet and return the high level 
        /// representation. See <see cref="RemoveColumn"/>
        /// </summary>
        /// <param name="columnnum">column number</param>
        /// <returns>High level <see cref="XSSFColumn"/> object representing a 
        /// column in the sheet</returns>
        public virtual IColumn CreateColumn(int columnnum)
        {
            CT_Col ctCol;
            XSSFColumn prev = _columns.TryGetValue(columnnum, out XSSFColumn column) ? column : null;
            if (prev != null)
            {
                // the Cells in an existing column are invalidated on-purpose, in
                // order to clean up correctly, we need to call the remove, so
                // things like ArrayFormulas and CalculationChain updates are
                // done correctly. We remove the cell this way as the internal
                // cell-list is changed by the remove call and thus would cause
                // ConcurrentModificationException otherwise
                while(prev.FirstCellNum != -1)
                {
                    prev.RemoveCell(prev.GetCell(prev.FirstCellNum));
                }

                ctCol = prev.GetCTCol();
                ctCol.Set(new CT_Col());
            }
            else
            {
                if(worksheet.cols.FirstOrDefault() is null)
                {
                    worksheet.AddNewCols();
                }

                if(_columns.Count == 0 || columnnum > GetLastKey(_columns.Keys))
                {
                    // we can append the new column at the end
                    ctCol = worksheet.cols.FirstOrDefault().AddNewCol();
                }
                else
                {
                    // get number of columns where column index < columnnum
                    // --> this tells us where our column should go
                    int idx = HeadMapCount(_columns.Keys, columnnum);

                    ctCol = worksheet.cols.FirstOrDefault().InsertNewCol(idx);
                }
            }

            ctCol.SetNumber((uint) columnnum + 1);

            XSSFColumn c = new XSSFColumn(ctCol, this) { ColumnNum = columnnum };

            _columns[columnnum] = c;
            return c;
        }

        /// <summary>
        /// Creates a split pane. Any existing freezepane or split pane is 
        /// overwritten.
        /// </summary>
        /// <param name="xSplitPos">Horizonatal position of split (in 1/20th 
        /// of a point).</param>
        /// <param name="ySplitPos">Vertical position of split (in 1/20th of 
        /// a point).</param>
        /// <param name="leftmostColumn">Left column visible in right pane.</param>
        /// <param name="topRow">Top row visible in bottom pane</param>
        /// <param name="activePane">Active pane.  One of: PANE_LOWER_RIGHT, 
        /// PANE_UPPER_RIGHT, PANE_LOWER_LEFT, PANE_UPPER_LEFT</param>
        public void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow,
            PanePosition activePane)
        {
            CreateFreezePane(xSplitPos, ySplitPos, leftmostColumn, topRow);
            GetPane().state = ST_PaneState.split;
            GetPane().activePane = (ST_Pane) activePane;
        }

        /// <summary>
        /// Returns cell comment for the specified row and column
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <returns>cell comment or <code>null</code> if not found</returns>
        [Obsolete(
            "deprecated as of 2015-11-23 (circa POI 3.14beta1). Use {@link #getCellComment(CellAddress)} instead.")]
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
            if(sheetComments == null)
            {
                return null;
            }

            int row = address.Row;
            int column = address.Column;

            CellAddress ref1 = new CellAddress(row, column);
            CT_Comment ctComment = sheetComments.GetCTComment(ref1);
            if(ctComment == null)
            {
                return null;
            }

            XSSFVMLDrawing vml = GetVMLDrawing(false);
            return new XSSFComment(sheetComments, ctComment,
                vml?.FindCommentShape(row, column));
        }

        /// <summary>
        /// Returns all cell comments on this sheet.
        /// </summary>
        /// <returns>return A Dictionary of each Comment in the sheet, keyed on
        /// the cell address where the comment is located.</returns>
        public Dictionary<CellAddress, IComment> GetCellComments()
        {
            if(sheetComments == null)
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
        /// <returns>return hyperlink if there is a hyperlink anchored at row, 
        /// column; otherwise returns null</returns>
        public IHyperlink GetHyperlink(int row, int column)
        {
            return GetHyperlink(new CellAddress(row, column));
        }

        /// <summary>
        /// Get a Hyperlink in this sheet located in a cell specified 
        /// by {code addr}
        /// </summary>
        /// <param name="addr">The address of the cell containing the 
        /// hyperlink</param>
        /// <returns>return hyperlink if there is a hyperlink anchored at 
        /// {@code addr}; otherwise returns {@code null}</returns>
        public IHyperlink GetHyperlink(CellAddress addr)
        {
            string ref1 = addr.FormatAsString();
            foreach(XSSFHyperlink hyperlink in hyperlinks)
            {
                if(hyperlink.CellRef.Equals(ref1))
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
            return hyperlinks.ToList<IHyperlink>();
        }

        /// <summary>
        /// Get the actual column width (in units of 1/256th of a character width)
        /// </summary>
        /// <param name="columnIndex">the column to set (0-based)</param>
        /// <returns>the width in units of 1/256th of a character width</returns>
        public double GetColumnWidth(int columnIndex)
        {
            IColumn col = GetColumn(columnIndex);

            double width = (col == null)
                ? DefaultColumnWidth
                : col.Width;
            return Math.Round(width * 256, 2);
        }

        /// <summary>
        /// Get the actual column width in pixels
        /// Please note, that this method works correctly only for workbooks
        /// with the default font size(Calibri 11pt for .xlsx).
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public double GetColumnWidthInPixels(int columnIndex)
        {
            double widthIn256 = GetColumnWidth(columnIndex);
            return widthIn256 / 256.0 * XSSFWorkbook.DEFAULT_CHARACTER_WIDTH;
        }

        /// <summary>
        /// Gets the size of the margin in inches.
        /// </summary>
        /// <param name="margin">which margin to get</param>
        /// <returns>the size of the margin</returns>
        /// <exception cref="ArgumentException"></exception>
        public double GetMargin(MarginType margin)
        {
            if(!worksheet.IsSetPageMargins())
            {
                return 0;
            }

            CT_PageMargins pageMargins = worksheet.pageMargins;
            switch(margin)
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
                    throw new ArgumentException(
                        "Unknown margin constant:  " + margin);
            }
        }

        /// <summary>
        /// Sets the size of the margin in inches.
        /// </summary>
        /// <param name="margin">which margin to get</param>
        /// <param name="size">the size of the margin</param>
        /// <exception cref="InvalidOperationException"></exception>
        public void SetMargin(MarginType margin, double size)
        {
            CT_PageMargins pageMargins =
                worksheet.IsSetPageMargins() ? worksheet.pageMargins : worksheet.AddNewPageMargins();
            switch(margin)
            {
                case MarginType.LeftMargin:
                    pageMargins.left = size;
                    break;
                case MarginType.RightMargin:
                    pageMargins.right = size;
                    break;
                case MarginType.TopMargin:
                    pageMargins.top = size;
                    break;
                case MarginType.BottomMargin:
                    pageMargins.bottom = size;
                    break;
                case MarginType.HeaderMargin:
                    pageMargins.header = size;
                    break;
                case MarginType.FooterMargin:
                    pageMargins.footer = size;
                    break;
                default:
                    throw new InvalidOperationException(
                        "Unknown margin constant:  " + margin);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <returns>the merged region at the specified index</returns>
        /// <exception cref="InvalidOperationException">if this worksheet 
        /// does not contain merged regions</exception>
        public CellRangeAddress GetMergedRegion(int index)
        {
            CT_MergeCells ctMergeCells = worksheet.mergeCells;
            if(ctMergeCells == null)
            {
                throw new InvalidOperationException("This worksheet does not contain merged regions");
            }

            CT_MergeCell ctMergeCell = ctMergeCells.GetMergeCellArray(index);

            if(ctMergeCell == null)
            {
                return null;
            }

            string ref1 = ctMergeCell.@ref;
            return CellRangeAddress.ValueOf(ref1);
        }

        public CellRangeAddress GetMergedRegion(CellRangeAddress mergedRegion)
        {
            if(worksheet.mergeCells == null || worksheet.mergeCells.mergeCell == null)
            {
                return null;
            }

            foreach(CT_MergeCell mc in worksheet.mergeCells.mergeCell)
            {
                if(mc != null && !string.IsNullOrEmpty(mc.@ref))
                {
                    CellRangeAddress range = CellRangeAddress.ValueOf(mc.@ref);
                    if(range.FirstColumn <= mergedRegion.FirstColumn
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

        /// <summary>
        /// Enables sheet protection and Sets the password for the sheet. 
        /// Also Sets some attributes on the { @link CT_SheetProtection } 
        /// that correspond to the default values used by Excel
        /// </summary>
        /// <param name="password">password to set for protection. Pass null 
        /// to remove protection</param>
        public void ProtectSheet(string password)
        {

            if(password != null)
            {
                CT_SheetProtection sheetProtection = worksheet.AddNewSheetProtection();
                SetSheetPassword(password, null); // defaults to xor password
                sheetProtection.sheet = true;
                sheetProtection.scenarios = true;
                sheetProtection.objects = true;
            }
            else
            {
                worksheet.UnsetSheetProtection();
            }
        }

        /// <summary>
        /// Sets the sheet password. 
        /// </summary>
        /// <param name="password"> if null, the password will be removed</param>
        /// <param name="hashAlgo">if null, the password will be set as XOR 
        /// password (Excel 2010 and earlier)otherwise the given algorithm is 
        /// used for calculating the hash password (Excel 2013)</param>
        public void SetSheetPassword(string password, HashAlgorithm hashAlgo)
        {
            if(password == null && !IsSheetProtectionEnabled())
            {
                return;
            }

            XSSFPasswordHelper.SetPassword(SafeGetProtectionField(),
                password, hashAlgo, null);
        }

        /// <summary>
        /// Validate the password against the stored hash, the hashing method 
        /// will be determined by the existing password attributes
        /// </summary>
        /// <param name="password"></param>
        /// <returns>true, if the hashes match (... though original password 
        /// may differ ...)</returns>
        public bool ValidateSheetPassword(string password)
        {
            if(!IsSheetProtectionEnabled())
            {
                return password == null;
            }

            return XSSFPasswordHelper.ValidatePassword(SafeGetProtectionField(),
                password, null);
        }

        /// <summary>
        /// Returns the logical row ( 0-based).  If you ask for a row that is 
        /// not defined you get a null.  This is to say row 4 represents the 
        /// fifth row on a sheet.
        /// </summary>
        /// <param name="rownum">row to get</param>
        /// <returns><see cref="XSSFRow"/> representing the rownumber or null 
        /// if its not defined on the sheet</returns>
        public IRow GetRow(int rownum)
        {
            if (_rows.TryGetValue(rownum, out XSSFRow row))
            {
                return row;
            }

            return null;
        }

        /// <summary>
        /// Returns the logical column ( 0-based).  If you ask for a column that is 
        /// not defined you get a null.  This is to say column 4 represents the 
        /// fifth column on a sheet.
        /// </summary>
        /// <param name="columnnum">column to get</param>
        /// <returns><see cref="XSSFColumn"/> representing the columnnumber or null 
        /// if its not defined on the sheet</returns>
        public IColumn GetColumn(int columnnum, bool createIfNull = false)
        {
            if (_columns.TryGetValue(columnnum, out XSSFColumn column))
            {
                return column;
            }

            if(createIfNull)
            {
                return CreateColumn(columnnum);
            }

            return null;
        }

        /// <summary>
        /// Group between (0 based) columns
        /// </summary>
        /// <param name="fromColumn"></param>
        /// <param name="toColumn"></param>
        public void GroupColumn(int fromColumn, int toColumn)
        {
            for(int i = fromColumn; i <= toColumn; i++)
            {
                IColumn column = GetColumn(i, true);
                column.OutlineLevel++;
            }

            SetSheetFormatPrOutlineLevelCol();
        }

        /// <summary>
        /// Determines if there is a page break at the indicated column
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public bool IsColumnBroken(int column)
        {
            int[] colBreaks = ColumnBreaks;
            for(int i = 0; i < colBreaks.Length; i++)
            {
                if(colBreaks[i] == column)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Get the hidden state for a given column.
        /// </summary>
        /// <param name="columnIndex">the column to set (0-based)</param>
        /// <returns>hidden - false if the column is visible</returns>
        public bool IsColumnHidden(int columnIndex)
        {
            IColumn col = GetColumn(columnIndex);
            return col != null && col.Hidden;
        }

        /// <summary>
        /// Tests if there is a page break at the indicated row
        /// </summary>
        /// <param name="row">index of the row to test</param>
        /// <returns>true if there is a page break at the indicated row</returns>
        public bool IsRowBroken(int row)
        {
            int[] rowBreaks = RowBreaks;
            for(int i = 0; i < rowBreaks.Length; i++)
            {
                if(rowBreaks[i] == row)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Sets a page break at the indicated row Breaks occur above the 
        /// specified row and left of the specified column inclusive. For 
        /// example, sheet.SetColumnBreak(2); breaks the sheet into two parts 
        /// with columns A,B,C in the first and D,E,... in the second. 
        /// Simuilar, sheet.SetRowBreak(2); breaks the sheet into two parts 
        /// with first three rows (rownum=1...3) in the first part and rows 
        /// starting with rownum=4 in the second.
        /// </summary>
        /// <param name="row">the row to break, inclusive</param>
        public void SetRowBreak(int row)
        {

            CT_PageBreak pgBreak = worksheet.IsSetRowBreaks()
                ? worksheet.rowBreaks
                : worksheet.AddNewRowBreaks();

            if(!IsRowBroken(row))
            {
                CT_Break brk = pgBreak.AddNewBrk();
                brk.id = (uint) row + 1; // this is id of the row element which is 1-based: <row r="1" ... >
                brk.man = true;
                brk.max = (uint) SpreadsheetVersion.EXCEL2007.LastColumnIndex; //end column of the break

                pgBreak.count = (uint) pgBreak.sizeOfBrkArray();
                pgBreak.manualBreakCount = (uint) pgBreak.sizeOfBrkArray();
            }
        }

        /// <summary>
        /// Removes a page break at the indicated column
        /// </summary>
        /// <param name="column"></param>
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void RemoveColumnBreak(int column)
        {
            if(!worksheet.IsSetColBreaks())
            {
                // no breaks
                return;
            }

            CT_PageBreak pgBreak = worksheet.colBreaks;
            List<CT_Break> brkArray = pgBreak.brk;
            for(int i = 0; i < brkArray.Count; i++)
            {
                if(brkArray[i].id == (column + 1))
                {
                    pgBreak.RemoveBrk(i);
                }
            }
        }

        /// <summary>
        /// Removes a merged region of cells (hence letting them free)
        /// </summary>
        /// <param name="index"></param>
        public void RemoveMergedRegion(int index)
        {
            CT_MergeCells ctMergeCells = worksheet.mergeCells;

            int size = ctMergeCells.sizeOfMergeCellArray();
            CT_MergeCell[] mergeCellsArray = new CT_MergeCell[size - 1];
            for(int i = 0; i < size; i++)
            {
                if(i < index)
                {
                    mergeCellsArray[i] = ctMergeCells.GetMergeCellArray(i);
                }
                else if(i > index)
                {
                    mergeCellsArray[i - 1] = ctMergeCells.GetMergeCellArray(i);
                }
            }

            if(mergeCellsArray.Length > 0)
            {
                ctMergeCells.SetMergeCellArray(mergeCellsArray);
            }
            else
            {
                worksheet.UnsetMergeCells();
            }
        }

        /// <summary>
        /// Removes a number of merged regions of cells (hence letting them 
        /// free) This method can be used to bulk-remove merged regions in a 
        /// way much faster than calling RemoveMergedRegion() for every single 
        /// merged region.
        /// </summary>
        /// <param name="indices">A Set of the regions to unmerge</param>
        public void RemoveMergedRegions(IList<int> indices)
        {
            if(!worksheet.IsSetMergeCells())
            {
                return;
            }

            CT_MergeCells ctMergeCells = worksheet.mergeCells;
            //TODO: The following codes are not same as poi, re-do it?.
            int size = ctMergeCells.sizeOfMergeCellArray();
            List<CT_MergeCell> newMergeCells =
                new List<CT_MergeCell>(ctMergeCells.sizeOfMergeCellArray());

            for(int i = 0, d = 0; i < size; i++)
            {
                if(!indices.Contains(i))
                {
                    //newMergeCells[d] = ctMergeCells.GetMergeCellArray(i);
                    newMergeCells.Add(ctMergeCells.GetMergeCellArray(i));
                    d++;
                }
            }

            if(ListIsEmpty(newMergeCells))
            {
                worksheet.UnsetMergeCells();
            }
            else
            {
                ctMergeCells.SetMergeCellArray(newMergeCells.ToArray());
            }
        }

        /// <summary>
        /// Remove a row from this sheet.  All cells Contained in the row are 
        /// Removed as well
        /// </summary>
        /// <param name="row">the row to Remove.</param>
        /// <exception cref="ArgumentException"></exception>
        public void RemoveRow(IRow row)
        {
            if(row.Sheet != this)
            {
                throw new ArgumentException("Specified row does not belong to" +
                                            " this sheet");
            }

            // collect cells into a temporary array to avoid ConcurrentModificationException
            List<XSSFCell> cellsToDelete = new List<XSSFCell>();
            foreach(ICell cell in row)
            {
                cellsToDelete.Add((XSSFCell) cell);
            }

            foreach(XSSFCell cell in cellsToDelete)
            {
                row.RemoveCell(cell);
            }

            int idx = _rows.Count(p => p.Key < row.RowNum); // _rows.headMap(row.getRowNum()).size();
            _rows.Remove(row.RowNum);
            worksheet.sheetData.RemoveRow(row.RowNum + 1); // Note that rows in worksheet.sheetData is 1-based.

            // also remove any comment located in that row
            if(sheetComments != null)
            {
                foreach(CellAddress ref1 in GetCellComments().Keys)
                {
                    if(ref1.Row == row.RowNum)
                    {
                        sheetComments.RemoveComment(ref1);
                    }
                }
            }
        }

        /// <summary>
        /// Remove a column from this sheet.  All cells Contained in the column are 
        /// Removed as well
        /// </summary>
        /// <param name="column">the column to Remove.</param>
        /// <exception cref="ArgumentException"></exception>
        public void RemoveColumn(IColumn column)
        {
            if(column == null)
            {
                throw new ArgumentException("Column can't be null");
            }

            if(column.Sheet != this)
            {
                throw new ArgumentException("Specified column does not belong to" +
                                            " this sheet");
            }

            // collect cells into a temporary array to avoid ConcurrentModificationException
            List<XSSFCell> cellsToDelete = new List<XSSFCell>();
            foreach(ICell cell in column)
            {
                cellsToDelete.Add((XSSFCell) cell);
            }

            foreach(XSSFCell cell in cellsToDelete)
            {
                column.RemoveCell(cell);
            }

            // also remove any comment located in that column
            if(sheetComments != null)
            {
                foreach(CellAddress ref1 in GetCellComments().Keys)
                {
                    if(ref1.Column == column.ColumnNum)
                    {
                        sheetComments.RemoveComment(ref1);
                    }
                }
            }

            if(hyperlinks != null)
            {
                foreach(XSSFHyperlink link in new List<XSSFHyperlink>(hyperlinks))
                {
                    CellReference ref1 = new CellReference(link.CellRef);
                    if(ref1.Col == column.ColumnNum)
                    {
                        hyperlinks.Remove(link);
                    }
                }
            }

            DestroyColumn(column);
        }

        /// <summary>
        /// Removes the page break at the indicated row
        /// </summary>
        /// <param name="row"></param>
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void RemoveRowBreak(int row)
        {
            if(!worksheet.IsSetRowBreaks())
            {
                return;
            }

            CT_PageBreak pgBreak = worksheet.rowBreaks;
            List<CT_Break> brkArray = pgBreak.brk;
            for(int i = 0; i < brkArray.Count; i++)
            {
                if(brkArray[i].id == (row + 1))
                {
                    pgBreak.RemoveBrk(i);
                }
            }
        }

        /// <summary>
        /// Sets a page break at the indicated column. Breaks occur above the 
        /// specified row and left of the specified column inclusive. For 
        /// example, sheet.SetColumnBreak(2); breaks the sheet into two parts 
        /// with columns A,B,C in the first and D,E,... in the second. 
        /// Simuilar, sheet.SetRowBreak(2); breaks the sheet into two parts 
        /// with first three rows (rownum=1...3) in the first part and rows 
        /// starting with rownum=4 in the second.
        /// </summary>
        /// <param name="column">the column to break, inclusive</param>
        public void SetColumnBreak(int column)
        {
            if(!IsColumnBroken(column))
            {
                CT_PageBreak pgBreak = worksheet.IsSetColBreaks() ? worksheet.colBreaks : worksheet.AddNewColBreaks();
                CT_Break brk = pgBreak.AddNewBrk();
                brk.id = (uint) column + 1; // this is id of the row element which is 1-based: <row r="1" ... >
                brk.man = true;
                brk.max = (uint) SpreadsheetVersion.EXCEL2007.LastRowIndex; //end row of the break

                pgBreak.count = (uint) pgBreak.sizeOfBrkArray();
                pgBreak.manualBreakCount = (uint) pgBreak.sizeOfBrkArray();
            }
        }

        public void SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            IColumn col = GetColumn(columnNumber);

            if(col == null)
            {
                return;
            }

            SortedDictionary<int, IColumn> group = GetAdjacentOutlineColumns(col);

            int lastColumnIndex = group.Keys.Max();

            foreach(KeyValuePair<int, IColumn> columnEntry in group)
            {
                if(columnEntry.Key == lastColumnIndex)
                {
                    columnEntry.Value.Collapsed = collapsed;
                }
                else
                {
                    columnEntry.Value.Hidden = collapsed;
                }
            }
        }

        /// <summary>
        /// Get the visibility state for a given column.
        /// </summary>
        /// <param name="columnIndex">the column to get (0-based)</param>
        /// <param name="hidden">the visiblity state of the column</param>
        public void SetColumnHidden(int columnIndex, bool hidden)
        {
            IColumn column = GetColumn(columnIndex, true);
            column.Hidden = hidden;
        }

        /// <summary>
        /// Set the width (in units of 1/256th of a character width)
        /// <para>
        /// The maximum column width for an individual cell is 255 
        /// characters. This value represents the number of characters that can
        /// be displayed in a cell that is formatted with the standard font 
        /// (first font in the workbook).
        /// </para>
        /// <para>
        /// Character width is defined as the maximum digit width of the 
        /// numbers 
        /// <code>
        /// 0, 1, 2, ... 9
        /// </code>
        /// as rendered using the default 
        /// font (first font in the workbook). Unless you are using a very 
        /// special font, the default character is '0' (zero), this is true for
        /// Arial (default font font in HSSF) and Calibri (default font in 
        /// XSSF)
        /// </para>
        /// <para>
        /// Please note, that the width set by this method includes 4 
        /// pixels of margin pAdding (two on each side), plus 1 pixel pAdding 
        /// for the gridlines (Section 3.3.1.12 of the OOXML spec). This 
        /// results is a slightly less value of visible characters than passed 
        /// to this method (approx. 1/2 of a character).
        /// </para>
        /// <para>
        /// To compute
        /// the actual number of visible characters, Excel uses the following 
        /// formula (Section 3.3.1.12 of the OOXML spec):
        /// </para>
        /// <code>
        /// width = TRuncate([{Number of Visible Characters} * 
        /// {Maximum Digit Width} + {5 pixel pAdding}]/{Maximum Digit Width}*256)/256
        /// </code>
        /// <para>
        /// Using the Calibri font as an example, the maximum digit width
        /// of 11 point font size is 7 pixels (at 96 dpi). If you set a column 
        /// width to be eight characters wide, e.g. 
        /// <code>
        /// SetColumnWidth(columnIndex, 8*256)
        /// </code>
        /// , then the actual value of visible characters (the value Shown in 
        /// Excel) is derived from the following equation: 
        /// <code>
        /// TRuncate([numChars*7+5]/7*256)/256 = 8;
        /// </code>
        /// which gives 
        /// <code>
        /// 7.29
        /// </code>.
        /// </para>
        /// </summary>
        /// <param name="columnIndex">the column to set (0-based)</param>
        /// <param name="width">the width in units of 1/256th of a character 
        /// width</param>
        /// <exception cref="ArgumentException">if width more than 255*256 (the
        /// maximum column width in Excel is 255 characters)</exception>
        public void SetColumnWidth(int columnIndex, double width)
        {
            if(width > 255 * 256)
            {
                throw new ArgumentException("The maximum column width for an " +
                                            "individual cell is 255 characters.");
            }

            IColumn column = GetColumn(columnIndex, true);
            column.Width = (double) width / 256;
        }

        public void SetDefaultColumnStyle(int column, ICellStyle style)
        {
            IColumn col = GetColumn(column, true);
            col.ColumnStyle = style;
        }

        /// <summary>
        /// group the row It is possible for collapsed to be false and yet 
        /// still have the rows in question hidden. This can be achieved by 
        /// having a lower outline level collapsed, thus hiding all the child 
        /// rows. Note that in this case, if the lowest level were expanded, 
        /// the middle level would remain collapsed.
        /// </summary>
        /// <param name="rowIndex">the row involved, 0 based</param>
        /// <param name="collapse">bool value for collapse</param>
        public void SetRowGroupCollapsed(int rowIndex, bool collapse)
        {
            if(collapse)
            {
                CollapseRow(rowIndex);
            }
            else
            {
                ExpandRow(rowIndex);
            }
        }

        /// <summary>
        /// Sets the zoom magnification for the sheet.  The zoom is expressed 
        /// as a fraction.  For example to express a zoom of 75% use 3 for the 
        /// numerator and 4 for the denominator.
        /// </summary>
        /// <param name="numerator">The numerator for the zoom 
        /// magnification.</param>
        /// <param name="denominator">The denominator for the zoom 
        /// magnification.</param>
        [Obsolete("deprecated 2015-11-23 (circa POI 3.14beta1). Use {@link #setZoom(int)} instead.")]
        public void SetZoom(int numerator, int denominator)
        {
            int zoom = 100 * numerator / denominator;
            SetZoom(zoom);
        }

        /// <summary>
        /// Window zoom magnification for current view representing percent 
        /// values. Valid values range from 10 to 400. Horizontal &amp; 
        /// Vertical scale toGether. For example:
        /// <code>
        /// 10 - 10%
        /// 20 - 20%
        /// ...
        /// 100 - 100%
        /// ...
        /// 400 - 400%
        /// </code>
        /// Current view can be Normal, Page Layout, or Page Break Preview.
        /// </summary>
        /// <param name="scale">window zoom magnification</param>
        /// <exception cref="ArgumentException">if scale is invalid</exception>
        public void SetZoom(int scale)
        {
            if(scale < 10 || scale > 400)
            {
                throw new ArgumentException("Valid scale values range from 10 to 400");
            }

            GetSheetTypeSheetView().zoomScale = (uint) scale;
        }

        /// <summary>
        /// copyRows rows from srcRows to this sheet starting at destStartRow 
        /// Additionally copies merged regions that are completely defined in 
        /// these rows (ie. merged 2 cells on a row to be shifted).
        /// </summary>
        /// <param name="srcRows">the rows to copy. Formulas will be offset by 
        /// the difference in the row number of the first row in srcRows and 
        /// destStartRow (even if srcRows are from a different sheet).</param>
        /// <param name="destStartRow">the row in this sheet to paste the first
        /// row of srcRows the remainder of srcRows will be pasted below 
        /// destStartRow per the cell copy policy</param>
        /// <param name="policy">is the cell copy policy, which can be used to 
        /// merge the source and destination when the source is blank, copy 
        /// styles only, paste as value, etc</param>
        /// <exception cref="ArgumentException"></exception>
        public void CopyRows(List<XSSFRow> srcRows, int destStartRow, CellCopyPolicy policy)
        {
            if(srcRows == null || srcRows.Count == 0)
            {
                throw new ArgumentException("No rows to copy");
            }

            IRow srcStartRow = srcRows[0];
            IRow srcEndRow = srcRows[srcRows.Count - 1];

            if(srcStartRow == null)
            {
                throw new ArgumentException("copyRows: First row cannot be null");
            }

            int srcStartRowNum = srcStartRow.RowNum;
            int srcEndRowNum = srcEndRow.RowNum;

            // check row numbers to make sure they are continuous and
            // increasing (monotonic) and srcRows does not contain null rows
            int size = srcRows.Count;
            for(int index = 1; index < size; index++)
            {
                IRow curRow = srcRows[index];
                if(curRow == null)
                {
                    throw new ArgumentException(
                        "srcRows may not contain null rows. Found null row at" +
                        " index " + index + ".");
                }
                else if(srcStartRow.Sheet.Workbook != curRow.Sheet.Workbook)
                {
                    throw new ArgumentException(
                        "All rows in srcRows must belong to the same sheet in" +
                        " the same workbook. Expected all rows from same " +
                        "workbook (" + srcStartRow.Sheet.Workbook + "). " +
                        "Got srcRows[" + index + "] from different workbook (" +
                        curRow.Sheet.Workbook + ").");
                }
                else if(srcStartRow.Sheet != curRow.Sheet)
                {
                    throw new ArgumentException(
                        "All rows in srcRows must belong to the same sheet. " +
                        "Expected all rows from " + srcStartRow.Sheet.SheetName +
                        ". Got srcRows[" + index + "] from " + curRow.Sheet.SheetName);
                }
            }

            // FIXME: is special behavior needed if srcRows and destRows belong to the same sheets and the regions overlap?

            CellCopyPolicy options = new CellCopyPolicy(policy) {
                // avoid O(N^2) performance scanning through all regions for
                // each row merged regions will be copied after all the rows
                // have been copied
                IsCopyMergedRegions = false
            };

            // FIXME: if srcRows contains gaps or null values, clear out those rows that will be overwritten
            // how will this work with merging (copy just values, leave cell styles in place?)

            int r = destStartRow;
            foreach(IRow srcRow in srcRows)
            {
                int destRowNum;
                if(policy.IsCondenseRows)
                {
                    destRowNum = r++;
                }
                else
                {
                    int shift = srcRow.RowNum - srcStartRowNum;
                    destRowNum = destStartRow + shift;
                }

                //removeRow(destRowNum); //this probably clears all external formula references to destRow, causing unwanted #REF! errors
                XSSFRow destRow = CreateRow(destRowNum) as XSSFRow;
                destRow.CopyRowFrom(srcRow, options);
            }

            // Only do additional copy operations here that cannot be done with
            // Row.copyFromRow(Row, options) reasons: operation needs to
            // interact with multiple rows or sheets.
            // Copy merged regions that are contained within the copy region
            if(policy.IsCopyMergedRegions)
            {
                // FIXME: is this something that rowShifter could be doing?
                int shift = destStartRow - srcStartRowNum;
                foreach(CellRangeAddress srcRegion in srcStartRow.Sheet.MergedRegions)
                {
                    if(srcStartRowNum <= srcRegion.FirstRow && srcRegion.LastRow <= srcEndRowNum)
                    {
                        // srcRegion is fully inside the copied rows
                        CellRangeAddress destRegion = srcRegion.Copy();
                        destRegion.FirstRow += shift;
                        destRegion.LastRow += shift;
                        AddMergedRegion(destRegion);
                    }
                }
            }
        }

        /// <summary>
        /// Copies rows between srcStartRow and srcEndRow to the same sheet, 
        /// starting at destStartRow Convenience function for 
        /// <see cref="CopyRows"/> Equivalent to 
        /// copyRows(getRows(srcStartRow, srcEndRow, false), destStartRow, cellCopyPolicy)
        /// </summary>
        /// <param name="srcStartRow">the index of the first row to copy the 
        /// cells from in this sheet</param>
        /// <param name="srcEndRow">the index of the last row to copy the cells
        /// from in this sheet</param>
        /// <param name="destStartRow">the index of the first row to copy the 
        /// cells to in this sheet</param>
        /// <param name="cellCopyPolicy">the policy to use to determine how 
        /// cells are copied</param>
        public void CopyRows(int srcStartRow, int srcEndRow, int destStartRow, CellCopyPolicy cellCopyPolicy)
        {
            List<XSSFRow> srcRows = GetRows(srcStartRow, srcEndRow, false);
            CopyRows(srcRows, destStartRow, cellCopyPolicy);
        }

        /// <summary>
        /// Shifts rows between startRow and endRow n number of rows. If you 
        /// use a negative number, it will shift rows up. Code ensures that 
        /// rows don't wrap around. 
        /// Calls ShiftRows(startRow, endRow, n, false, false);
        /// <para>
        /// Additionally Shifts merged regions that are completely defined in 
        /// these rows (ie. merged 2 cells on a row to be Shifted).
        /// </para>
        /// </summary>
        /// <param name="startRow">the row to start Shifting</param>
        /// <param name="endRow">the row to end Shifting</param>
        /// <param name="n">the number of rows to shift</param>
        public void ShiftRows(int startRow, int endRow, int n)
        {
            ShiftRows(startRow, endRow, n, false, false);
        }

        /// <summary>
        /// Shifts columns between startColumn and endColumn n number of columns. If you 
        /// use a negative number, it will shift columns left. Code ensures that 
        /// columns don't wrap around. 
        /// Calls ShiftColumns(startColumn, endColumn, n, false, false);
        /// <para>
        /// Additionally Shifts merged regions that are completely defined in 
        /// these columns (ie. merged 2 cells on a column to be Shifted).
        /// </para>
        /// </summary>
        /// <param name="startColumn">the column to start Shifting</param>
        /// <param name="endColumn">the column to end Shifting</param>
        /// <param name="n">the number of column to shift</param>
        public void ShiftColumns(int startColumn, int endColumn, int n)
        {
            ShiftColumns(startColumn, endColumn, n, false, false);
        }

        /// <summary>
        /// Shifts rows between startRow and endRow n number of rows. If you 
        /// use a negative number, it will shift rows up. Code ensures that 
        /// rows don't wrap around
        /// <para>
        /// Additionally Shifts merged regions thatare completely defined in 
        /// these rows (ie. merged 2 cells on a row to be Shifted).
        /// </para>
        /// </summary>
        /// <param name="startRow">the row to start Shifting</param>
        /// <param name="endRow">the row to end Shifting</param>
        /// <param name="n">the number of rows to shift</param>
        /// <param name="copyRowHeight">whether to copy the row height during 
        /// the shift</param>
        /// <param name="resetOriginalRowHeight">whether to set the original 
        /// row's height to the default</param>
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight)
        {
            int sheetIndex = Workbook.GetSheetIndex(this);
            string sheetName = Workbook.GetSheetName(sheetIndex);
            FormulaShifter shifter = FormulaShifter.CreateForRowShift(
                sheetIndex, sheetName, startRow, endRow, n, SpreadsheetVersion.EXCEL2007);
            RemoveOverwrittenRows(startRow, endRow, n);
            ShiftCommentsAndRows(startRow, endRow, n, copyRowHeight);

            XSSFRowShifter rowShifter = new XSSFRowShifter(this);
            rowShifter.ShiftMergedRegions(startRow, endRow, n);
            rowShifter.UpdateNamedRanges(shifter);
            rowShifter.UpdateFormulas(shifter);
            rowShifter.UpdateConditionalFormatting(shifter);
            rowShifter.UpdateHyperlinks(shifter);
            RebuildRows();
        }

        /// <summary>
        /// Shifts columns between startColumn and endColumn n number of columns. If you 
        /// use a negative number, it will shift columns left. Code ensures that 
        /// columns don't wrap around
        /// <para>
        /// Additionally Shifts merged regions thatare completely defined in 
        /// these columns (ie. merged 2 cells on a column to be Shifted).
        /// </para>
        /// </summary>
        /// <param name="startColumn">the column to start Shifting</param>
        /// <param name="endColumn">the column to end Shifting</param>
        /// <param name="n">the number of columns to shift</param>
        /// <param name="copyColumnWidth">whether to copy the column width during 
        /// the shift</param>
        /// <param name="resetOriginalColumnWidth">whether to set the original 
        /// column's width to the default</param>
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public void ShiftColumns(
            int startColumn,
            int endColumn,
            int n,
            bool copyColumnWidth,
            bool resetOriginalColumnWidth)
        {
            int sheetIndex = Workbook.GetSheetIndex(this);
            string sheetName = Workbook.GetSheetName(sheetIndex);
            FormulaShifter shifter = FormulaShifter.CreateForColumnShift(
                sheetIndex, sheetName, startColumn, endColumn, n, SpreadsheetVersion.EXCEL2007);
            RemoveOverwrittenColumns(startColumn, endColumn, n);
            ShiftCommentsAndColumns(startColumn, endColumn, n, copyColumnWidth);

            XSSFColumnShifter columnShifter = new XSSFColumnShifter(this);
            columnShifter.ShiftMergedRegions(startColumn, endColumn, n);
            columnShifter.UpdateNamedRanges(shifter);
            columnShifter.UpdateFormulas(shifter);
            columnShifter.UpdateConditionalFormatting(shifter);
            columnShifter.UpdateHyperlinks(shifter);
            RebuildColumns();
        }

        /// <summary>
        /// Returns the CellStyle that applies to the given (0 based) column, 
        /// or null if no style has been set for that column
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public ICellStyle GetColumnStyle(int column)
        {
            IColumn col = GetColumn(column);

            if(col != null)
            {
                return col.ColumnStyle;
            }

            return Workbook.GetCellStyleAt(0);
        }

        /// <summary>
        /// Tie a range of cell toGether so that they can be collapsed 
        /// or expanded
        /// </summary>
        /// <param name="fromRow">start row (0-based)</param>
        /// <param name="toRow">end row (0-based)</param>
        public void GroupRow(int fromRow, int toRow)
        {
            for(int i = fromRow; i <= toRow; i++)
            {
                XSSFRow xrow = (XSSFRow) GetRow(i);
                if(xrow == null)
                {
                    xrow = (XSSFRow) CreateRow(i);
                }

                CT_Row ctrow = xrow.GetCTRow();
                short outlineLevel = ctrow.outlineLevel;
                ctrow.outlineLevel = (byte) (outlineLevel + 1);
            }

            SetSheetFormatPrOutlineLevelRow();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row">the zero based row index to find from</param>
        /// <returns></returns>
        public int FindEndOfRowOutlineGroup(int row)
        {
            int level = ((XSSFRow) GetRow(row)).GetCTRow().outlineLevel;
            int currentRow;
            int lastRowNum = LastRowNum;
            for(currentRow = row; currentRow < lastRowNum; currentRow++)
            {
                if(GetRow(currentRow) == null
                   || ((XSSFRow) GetRow(currentRow)).GetCTRow().outlineLevel < level)
                {
                    break;
                }
            }

            return currentRow;
        }

        public void UngroupColumn(int fromColumn, int toColumn)
        {
            for(int index = fromColumn; index <= toColumn; index++)
            {
                IColumn col = GetColumn(index);
                if(col != null)
                {
                    int outlineLevel = col.OutlineLevel;
                    col.OutlineLevel = outlineLevel - 1 < 0 ? 0 : outlineLevel - 1;

                    if(col.OutlineLevel == 0
                       && col.FirstCellNum == -1
                       && (col.ColumnStyle == null
                           || col.ColumnStyle.Index == Workbook.GetCellStyleAt(0).Index))
                    {
                        DestroyColumn(col);
                    }
                }
            }

            SetSheetFormatPrOutlineLevelCol();
        }

        /// <summary>
        /// Ungroup a range of rows that were previously groupped
        /// </summary>
        /// <param name="fromRow">start row (0-based)</param>
        /// <param name="toRow">end row (0-based)</param>
        public void UngroupRow(int fromRow, int toRow)
        {
            for(int i = fromRow; i <= toRow; i++)
            {
                XSSFRow xrow = (XSSFRow) GetRow(i);
                if(xrow != null)
                {
                    CT_Row ctrow = xrow.GetCTRow();
                    short outlinelevel = ctrow.outlineLevel;
                    ctrow.outlineLevel = (byte) (outlinelevel - 1);
                    // remove a row only if the row has no cell and if the
                    // outline level is 0
                    if(ctrow.outlineLevel == 0 && xrow.FirstCellNum == -1)
                    {
                        RemoveRow(xrow);
                    }
                }
            }

            SetSheetFormatPrOutlineLevelRow();
        }

        /// <summary>
        /// Register a hyperlink in the collection of hyperlinks on this sheet
        /// </summary>
        /// <param name="hyperlink">the link to add</param>
        public void AddHyperlink(XSSFHyperlink hyperlink)
        {
            hyperlinks.Add(hyperlink);
        }

        /// <summary>
        /// Removes a hyperlink in the collection of hyperlinks on this sheet
        /// </summary>
        /// <param name="row">row index</param>
        /// <param name="column">column index</param>
        public void RemoveHyperlink(int row, int column)
        {
            // CTHyperlinks is regenerated from scratch when writing out the
            // spreadsheet so don't worry about maintaining hyperlinks and
            // CTHyperlinks in parallel. only maintain hyperlinks
            string ref1 = new CellReference(row, column).FormatAsString();
            for(int index = 0; index < hyperlinks.Count; index++)
            {
                XSSFHyperlink hyperlink = hyperlinks[index];
                if(hyperlink.CellRef.Equals(ref1))
                {
                    if(worksheet != null
                       && worksheet.hyperlinks != null
                       && worksheet.hyperlinks.hyperlink != null
                       && worksheet.hyperlinks.hyperlink.Contains(hyperlink.GetCTHyperlink()))
                    {
                        worksheet.hyperlinks.hyperlink.Remove(hyperlink.GetCTHyperlink());
                    }

                    hyperlinks.RemoveAt(index);
                    return;
                }
            }
        }

        [Obsolete("deprecated 3.14beta2 (circa 2015-12-05). Use {@link #setActiveCell(CellAddress)} instead.")]
        public void SetActiveCell(string cellref)
        {
            CT_Selection ctsel = GetSheetTypeSelection();
            ctsel.activeCell = cellref;
            ctsel.SetSqref(new string[] { cellref });
        }

        /// <summary>
        /// Enable sheet protection
        /// </summary>
        public void EnableLocking()
        {
            SafeGetProtectionField().sheet = true;
        }

        /// <summary>
        /// Disable sheet protection
        /// </summary>
        public void DisableLocking()
        {
            SafeGetProtectionField().sheet = false;
        }

        /// <summary>
        /// Enable or disable Autofilters locking. This does not modify sheet 
        /// protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockAutoFilter(bool enabled)
        {
            SafeGetProtectionField().autoFilter = enabled;
        }

        /// <summary>
        /// Enable or disable Deleting columns locking. This does not modify 
        /// sheet protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockDeleteColumns(bool enabled)
        {
            SafeGetProtectionField().deleteColumns = enabled;
        }

        /// <summary>
        /// Enable or disable Deleting rows locking. This does not modify 
        /// sheet protection status. To enforce this un-/locking, call 
        /// 
        /// </summary>
        /// <param name="enabled"></param>
        public void LockDeleteRows(bool enabled)
        {
            SafeGetProtectionField().deleteRows = enabled;
        }

        /// <summary>
        /// Enable or disable Formatting cells locking. This does not modify 
        /// sheet protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockFormatCells(bool enabled)
        {
            SafeGetProtectionField().formatCells = enabled;
        }

        /// <summary>
        /// Enable or disable Formatting columns locking. This does not modify 
        /// sheet protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockFormatColumns(bool enabled)
        {
            SafeGetProtectionField().formatColumns = enabled;
        }

        /// <summary>
        /// Enable or disable Formatting rows locking. This does not modify 
        /// sheet protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockFormatRows(bool enabled)
        {
            SafeGetProtectionField().formatRows = enabled;
        }

        /// <summary>
        /// Enable or disable Inserting columns locking. This does not modify 
        /// sheet protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockInsertColumns(bool enabled)
        {
            SafeGetProtectionField().insertColumns = enabled;
        }

        /// <summary>
        /// Enable or disable Inserting hyperlinks locking. This does not 
        /// modify sheet protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockInsertHyperlinks(bool enabled)
        {
            SafeGetProtectionField().insertHyperlinks = enabled;
        }

        /// <summary>
        /// Enable or disable Inserting rows locking. This does not modify 
        /// sheet protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockInsertRows(bool enabled)
        {
            SafeGetProtectionField().insertRows = enabled;
        }

        /// <summary>
        /// Enable or disable Pivot Tables locking. This does not modify sheet 
        /// protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockPivotTables(bool enabled)
        {
            SafeGetProtectionField().pivotTables = enabled;
        }

        /// <summary>
        /// Enable or disable Sort locking. This does not modify sheet 
        /// protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockSort(bool enabled)
        {
            SafeGetProtectionField().sort = enabled;
        }

        /// <summary>
        /// Enable or disable Objects locking. This does not modify sheet 
        /// protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockObjects(bool enabled)
        {
            SafeGetProtectionField().objects = enabled;
        }

        /// <summary>
        /// Enable or disable Scenarios locking. This does not modify sheet 
        /// protection status. To enforce this un-/locking, call 
        /// <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockScenarios(bool enabled)
        {
            SafeGetProtectionField().scenarios = enabled;
        }

        /// <summary>
        /// Enable or disable Selection of locked cells locking. This does not 
        /// modify sheet protection status. To enforce this un-/locking, call 
        /// 
        /// </summary>
        /// <param name="enabled"></param>
        public void LockSelectLockedCells(bool enabled)
        {
            SafeGetProtectionField().selectLockedCells = enabled;
        }

        /// <summary>
        /// Enable or disable Selection of unlocked cells locking. This does 
        /// not modify sheet protection status. To enforce this un-/locking, 
        /// call <see cref="DisableLocking"/> or <see cref="EnableLocking"/>
        /// </summary>
        /// <param name="enabled"></param>
        public void LockSelectUnlockedCells(bool enabled)
        {
            SafeGetProtectionField().selectUnlockedCells = enabled;
        }

        public ICellRange<ICell> SetArrayFormula(string formula, CellRangeAddress range)
        {

            ICellRange<ICell> cr = GetCellRange(range);

            ICell mainArrayFormulaCell = cr.TopLeftCell;
            ((XSSFCell) mainArrayFormulaCell).SetCellArrayFormula(formula, range);
            arrayFormulas.Add(range);
            return cr;
        }

        public ICellRange<ICell> RemoveArrayFormula(ICell cell)
        {
            if(cell.Sheet != this)
            {
                throw new ArgumentException("Specified cell does not belong " +
                                            "to this sheet.");
            }

            foreach(CellRangeAddress range in arrayFormulas)
            {
                if(range.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    arrayFormulas.Remove(range);
                    ICellRange<ICell> cr = GetCellRange(range);
                    foreach(ICell c in cr)
                    {
                        c.SetCellType(CellType.Blank);
                    }

                    return cr;
                }
            }

            string ref1 = ((XSSFCell) cell).GetCTCell().r;
            throw new ArgumentException(
                "Cell " + ref1 + " is not part of an array formula.");
        }

        public IDataValidationHelper GetDataValidationHelper()
        {
            return dataValidationHelper;
        }

        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public List<IDataValidation> GetDataValidations()
        {
            List<IDataValidation> xssfValidations = new List<IDataValidation>();
            CT_DataValidations dataValidations = worksheet.dataValidations;
            if(dataValidations != null && dataValidations.count > 0)
            {
                foreach(CT_DataValidation ctDataValidation in dataValidations.dataValidation)
                {
                    CellRangeAddressList addressList = new CellRangeAddressList();

                    string[] regions = ctDataValidation.sqref.Split(new char[] { ' ' });
                    for(int i = 0; i < regions.Length; i++)
                    {
                        if(regions[i].Length == 0)
                        {
                            continue;
                        }

                        string[] parts = regions[i].Split(new char[] { ':' });
                        CellReference begin = new CellReference(parts[0]);
                        CellReference end = parts.Length > 1
                            ? new CellReference(parts[1])
                            : begin;
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(
                            begin.Row,
                            end.Row,
                            begin.Col,
                            end.Col);
                        addressList.AddCellRangeAddress(cellRangeAddress);
                    }

                    XSSFDataValidation xssfDataValidation =
                        new XSSFDataValidation(addressList, ctDataValidation);
                    xssfValidations.Add(xssfDataValidation);
                }
            }

            return xssfValidations;
        }

        public void AddValidationData(IDataValidation dataValidation)
        {
            XSSFDataValidation xssfDataValidation = (XSSFDataValidation) dataValidation;
            CT_DataValidations dataValidations = worksheet.dataValidations;
            if(dataValidations == null)
            {
                dataValidations = worksheet.AddNewDataValidations();
            }

            int currentCount = dataValidations.sizeOfDataValidationArray();
            CT_DataValidation newval = dataValidations.AddNewDataValidation();
            newval.Set(xssfDataValidation.GetCTDataValidation());
            dataValidations.count = (uint) currentCount + 1;
        }

        public void RemoveDataValidation(IDataValidation dataValidation)
        {
            XSSFDataValidation xssfDataValidation = (XSSFDataValidation) dataValidation;
            CT_DataValidations dataValidations = worksheet.dataValidations;

            if(dataValidations is null)
            {
                return;
            }

            int currentCount = dataValidations.sizeOfDataValidationArray();

            for(int i = 0; i < currentCount; i++)
            {
                CT_DataValidation ctDataValidation = dataValidations.dataValidation[i];

                if(ctDataValidation.Equals(xssfDataValidation.GetCTDataValidation()))
                {
                    dataValidations.dataValidation.RemoveAt(i);
                    dataValidations.count = (uint) currentCount - 1;
                    return;
                }
            }
        }

        public IAutoFilter SetAutoFilter(CellRangeAddress range)
        {
            CT_AutoFilter af = worksheet.autoFilter;
            if(af == null)
            {
                af = worksheet.AddNewAutoFilter();
            }

            CellRangeAddress norm = new CellRangeAddress(range.FirstRow, range.LastRow,
                range.FirstColumn, range.LastColumn);
            string ref1 = norm.FormatAsString();
            af.@ref = ref1;

            XSSFWorkbook wb = (XSSFWorkbook) Workbook;
            int sheetIndex = Workbook.GetSheetIndex(this);
            XSSFName name = wb.GetBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
            if(name == null)
            {
                name = wb.CreateBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
            }

            name.GetCTName().hidden = true;
            CellReference r1 = new CellReference(SheetName, range.FirstRow, range.FirstColumn, true, true);
            CellReference r2 = new CellReference(null, range.LastRow, range.LastColumn, true, true);
            string fmla = r1.FormatAsString() + ":" + r2.FormatAsString();
            name.RefersToFormula = fmla;

            return new XSSFAutoFilter(this);
        }

        /// <summary>
        /// Creates a new Table, and associates it with this Sheet
        /// </summary>
        /// <returns></returns>
        public XSSFTable CreateTable()
        {
            if(!worksheet.IsSetTableParts())
            {
                worksheet.AddNewTableParts();
            }

            CT_TableParts tblParts = worksheet.tableParts;
            CT_TablePart tbl = tblParts.AddNewTablePart();

            // Table numbers need to be unique in the file, not just
            //  unique within the sheet. Find the next one
            int tableNumber = GetPackagePart().Package
                .GetPartsByContentType(XSSFRelation.TABLE.ContentType).Count + 1;
            RelationPart rp = CreateRelationship(
                XSSFRelation.TABLE,
                XSSFFactory.GetInstance(),
                tableNumber,
                false);
            XSSFTable table = rp.DocumentPart as XSSFTable;
            tbl.id = rp.Relationship.Id;

            tables[tbl.id] = table;

            return table;
        }

        /// <summary>
        /// Returns any tables associated with this Sheet
        /// </summary>
        /// <returns></returns>
        public List<XSSFTable> GetTables()
        {
            List<XSSFTable> tableList = new List<XSSFTable>(
                tables.Values
            );
            return tableList;
        }

        /// <summary>
        /// Set background color of the sheet tab
        /// </summary>
        /// <param name="colorIndex">the indexed color to set, must be a 
        /// constant from <see cref="IndexedColors"/></param>
        [Obsolete("deprecated 3.15-beta2. Removed in 3.17. Use {@link #setTabColor(XSSFColor)}.")]
        public void SetTabColor(int colorIndex)
        {
            CT_SheetPr pr = worksheet.sheetPr;
            if(pr == null)
            {
                pr = worksheet.AddNewSheetPr();
            }

            CT_Color color = new CT_Color { indexed = (uint) colorIndex };
            pr.tabColor = color;
        }

        #region ISheet Members

        public IDrawing DrawingPatriarch
        {
            get
            {
                if(drawing == null)
                {
                    OpenXmlFormats.Spreadsheet.CT_Drawing ctDrawing = GetCTDrawing();
                    if(ctDrawing == null)
                    {
                        return null;
                    }

                    foreach(RelationPart rp in RelationParts)
                    {
                        POIXMLDocumentPart p = rp.DocumentPart;
                        if(p is XSSFDrawing dr)
                        {
                            string drId = rp.Relationship.Id;
                            if(drId.Equals(ctDrawing.id))
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
            if(worksheet.mergeCells == null || worksheet.mergeCells.mergeCell == null)
            {
                return false;
            }

            foreach(CT_MergeCell mc in worksheet.mergeCells.mergeCell)
            {
                if(!string.IsNullOrEmpty(mc.@ref))
                {
                    CellRangeAddress range = CellRangeAddress.ValueOf(mc.@ref);
                    if(range.FirstColumn <= mergedRegion.FirstColumn
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
            IsSelected = value;
        }

        public void SetActiveCellRange(List<CellRangeAddress8Bit> cellranges, int activeRange, int activeRow,
            int activeColumn)
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
                CT_SheetView view = GetDefaultSheetView();
                return view != null && view.rightToLeft;
            }
            set
            {
                CT_SheetView view = GetDefaultSheetView();
                view.rightToLeft = value;
            }
        }

        #endregion

        public IRow CopyRow(int sourceIndex, int targetIndex)
        {
            return SheetUtil.CopyRow(this, sourceIndex, targetIndex);
        }

        /// <summary>
        /// Copy the source column to the target column. If the target column 
        /// exists, the new copied column will be inserted before the 
        /// existing one
        /// </summary>
        /// <param name="sourceIndex">source index</param>
        /// <param name="targetIndex">target index</param>
        /// <returns>the new copied column object</returns>
        public IColumn CopyColumn(int sourceIndex, int targetIndex)
        {
            if(sourceIndex == targetIndex)
            {
                throw new ArgumentException(
                    "sourceIndex and targetIndex cannot be same");
            }

            // Get the source / new column
            IColumn newColumn = GetColumn(targetIndex);
            IColumn sourceColumn = GetColumn(sourceIndex);

            // If the column exist in destination, push right all columns
            // by 1 else create a new column
            if(newColumn != null)
            {
                ShiftColumns(targetIndex, LastColumnNum, 1);
            }

            newColumn = CreateColumn(targetIndex);
            newColumn.Width = sourceColumn.Width; //copy column width

            // Loop through source cells to add to new column
            for(int i = sourceColumn.FirstCellNum; i < sourceColumn.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                ICell oldCell = sourceColumn.GetCell(i);

                // If the old cell is null jump to next cell
                if(oldCell == null)
                {
                    continue;
                }

                ICell newCell = newColumn.CreateCell(i);

                if(oldCell.CellStyle != null)
                {
                    // apply style from old cell to new cell 
                    newCell.CellStyle = oldCell.CellStyle;
                }

                if(oldCell.CellComment != null)
                {
                    CopyComment(oldCell, newCell);
                }

                // If there is a cell hyperlink, copy
                if(oldCell.Hyperlink != null)
                {
                    newCell.Hyperlink = oldCell.Hyperlink;
                }

                // Set the cell data type
                newCell.SetCellType(oldCell.CellType);

                // Set the cell data value
                switch(oldCell.CellType)
                {
                    case CellType.Blank:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                    case CellType.Boolean:
                        newCell.SetCellValue(oldCell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                        break;
                    case CellType.Formula:
                        newCell.SetCellFormula(oldCell.CellFormula);
                        break;
                    case CellType.Numeric:
                        newCell.SetCellValue(oldCell.NumericCellValue);
                        break;
                    case CellType.String:
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                        break;
                }
            }

            // If there are are any merged regions in the source column,
            // copy to new column
            for(int i = 0; i < NumMergedRegions; i++)
            {
                CellRangeAddress cellRangeAddress = GetMergedRegion(i);
                if(cellRangeAddress != null
                   && cellRangeAddress.FirstColumn == sourceColumn.ColumnNum)
                {
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(
                        cellRangeAddress.FirstRow,
                        cellRangeAddress.LastRow,
                        newColumn.ColumnNum,
                        newColumn.ColumnNum + cellRangeAddress.LastColumn - cellRangeAddress.FirstColumn);
                    AddMergedRegion(newCellRangeAddress);
                }
            }

            return newColumn;
        }

        /// <summary>
        /// This method will destroy the List <see cref="IColumn"/> object
        /// without affecting the cells, formulas, styles, etc contained in the
        /// sheet. It isuseful for when you've just created an IColumn to do
        /// some manipulation on and don't need it anymore, and don't want to
        /// leave it around in the sheet.</summary>
        /// <param name="column"><see cref="IColumn"/> to destroy</param>
        /// <exception cref="ArgumentException">If <paramref name="column"/> is
        /// null or if it doesn't belong to this <see cref="XSSFSheet"/></exception>
        public void DestroyColumns(List<IColumn> columnsToDestroy)
        {
            foreach(IColumn column in columnsToDestroy)
            {
                DestroyColumn(column);
            }
        }

        /// <summary>
        /// This method will destroy the <see cref="IColumn"/> object without
        /// affecting the cells, formulas, styles, etc contained in the sheet.
        /// It isuseful for when you've just created an IColumn to do some
        /// manipulation on and don't need it anymore, and don't want to leave
        /// it around in the sheet.
        /// </summary>
        /// <param name="column"><see cref="IColumn"/> to destroy</param>
        /// <exception cref="ArgumentException">If <paramref name="column"/> is
        /// null or if it doesn't belong to this <see cref="XSSFSheet"/></exception>
        public void DestroyColumn(IColumn column)
        {
            if(column == null)
            {
                throw new ArgumentException("Column can't be null");
            }

            if(column.Sheet != this)
            {
                throw new ArgumentException("Specified column does not belong to" +
                                            " this sheet");
            }

            _columns.Remove(column.ColumnNum);

            CT_Cols ctCols = worksheet.cols.FirstOrDefault();
            CT_Col ctCol = ((XSSFColumn) column).GetCTCol();
            int colIndex = GetIndexOfColumn(ctCols, ctCol);

            ctCols.RemoveCol(colIndex); // Note that columns in worksheet.sheetData is 1-based.
        }

        public void ShowInPane(int toprow, int leftcol)
        {
            CellReference cellReference = new CellReference(toprow, leftcol);
            string cellRef = cellReference.FormatAsString();
            Pane.topLeftCell = cellRef;
        }

        public ISheet CopySheet(string Name)
        {
            return CopySheet(Name, true);
        }

        public ISheet CopySheet(string name, bool copyStyle)
        {
            string clonedName = SheetUtil.GetUniqueSheetName(Workbook, name);
            XSSFSheet clonedSheet = (XSSFSheet) Workbook.CreateSheet(clonedName);

            try
            {
                using(MemoryStream ms = RecyclableMemory.GetStream())
                {
                    Write(ms, true);
                    ms.Position = 0;
                    clonedSheet.Read(ms);
                }
            }
            catch(IOException e)
            {
                throw new POIXMLException("Failed to clone sheet", e);
            }

            CT_Worksheet ct = clonedSheet.GetCTWorksheet();
            if(ct.IsSetLegacyDrawing())
            {
                logger.Log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
                ct.UnsetLegacyDrawing();
            }

            clonedSheet.IsSelected = false;

            // copy sheet's relations
            IList<POIXMLDocumentPart> rels = GetRelations();
            // if the sheet being cloned has a drawing then remember it and re-create too
            XSSFDrawing dg = null;
            foreach(POIXMLDocumentPart r in rels)
            {
                // do not copy the drawing relationship, it will be re-created
                if(r is XSSFDrawing dr)
                {
                    dg = dr;
                    continue;
                }

                //skip printerSettings.bin part
                if(r.GetPackagePart().PartName.Name.StartsWith("/xl/printerSettings/printerSettings"))
                {
                    continue;
                }

                PackageRelationship rel = r.GetPackageRelationship();
                clonedSheet.GetPackagePart().AddRelationship(
                    rel.TargetUri, (TargetMode) rel.TargetMode, rel.RelationshipType);
                clonedSheet.AddRelation(rel.Id, r);
            }

            // copy hyperlinks
            clonedSheet.hyperlinks = new List<XSSFHyperlink>(hyperlinks);

            // clone the sheet drawing along with its relationships
            if(dg != null)
            {
                if(ct.IsSetDrawing())
                {
                    // unset the existing reference to the drawing, so that
                    // subsequent call of clonedSheet.createDrawingPatriarch()
                    // will create a new one
                    ct.UnsetDrawing();
                }

                XSSFDrawing clonedDg = clonedSheet.CreateDrawingPatriarch() as XSSFDrawing;
                // copy drawing contents
                clonedDg.GetCTDrawing().Set(dg.GetCTDrawing());

                clonedDg = clonedSheet.CreateDrawingPatriarch() as XSSFDrawing;

                // Clone drawing relations
                IList<POIXMLDocumentPart> srcRels = dg.GetRelations();
                foreach(POIXMLDocumentPart rel in srcRels)
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

        public void CopyTo(IWorkbook dest, string name, bool copyStyle, bool keepFormulas)
        {
            StylesTable styles = ((XSSFWorkbook) dest).GetStylesSource();
            if(copyStyle && Workbook.NumberOfFonts > 0)
            {
                foreach(XSSFFont font in ((XSSFWorkbook) Workbook).GetStylesSource().GetFonts())
                {
                    styles.PutFont(font);
                }
            }

            XSSFSheet newSheet = (XSSFSheet) dest.CreateSheet(name);
            newSheet.sheet.state = sheet.state;
            IDictionary<int, ICellStyle> styleMap = copyStyle ? new Dictionary<int, ICellStyle>() : null;
            for(int i = FirstRowNum; i <= LastRowNum; i++)
            {
                XSSFRow srcRow = (XSSFRow) GetRow(i);
                XSSFRow destRow = (XSSFRow) newSheet.CreateRow(i);
                if(srcRow != null)
                {
                    // avoid O(N^2) performance scanning through all regions for each row
                    // merged regions will be copied after all the rows have been copied
                    CopyRow(this, newSheet, srcRow, destRow, styleMap, keepFormulas, keepMergedRegion: false);
                }
            }

            // Copying merged regions
            foreach(var srcRegion in this.MergedRegions)
            {
                var destRegion = srcRegion.Copy();
                newSheet.AddMergedRegion(destRegion);
            }

            List<CT_Cols> srcCols = worksheet.GetColsList();
            List<CT_Cols> dstCols = newSheet.worksheet.GetColsList();
            dstCols.Clear(); //Should already be empty since this is a new sheet.
            foreach(CT_Cols srcCol in srcCols)
            {
                CT_Cols dstCol = new CT_Cols();
                foreach(CT_Col column in srcCol.col)
                {
                    dstCol.col.Add(column.Copy());
                }

                dstCols.Add(dstCol);
            }

            newSheet.ForceFormulaRecalculation = true;
            newSheet.PrintSetup.Landscape = PrintSetup.Landscape;
            newSheet.PrintSetup.HResolution = PrintSetup.HResolution;
            newSheet.PrintSetup.VResolution = PrintSetup.VResolution;
            newSheet.SetMargin(
                MarginType.LeftMargin,
                GetMargin(MarginType.LeftMargin));
            newSheet.SetMargin(
                MarginType.RightMargin,
                GetMargin(MarginType.RightMargin));
            newSheet.SetMargin(
                MarginType.TopMargin,
                GetMargin(MarginType.TopMargin));
            newSheet.SetMargin(
                MarginType.BottomMargin,
                GetMargin(MarginType.BottomMargin));
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
            if(worksheet.IsSetSheetPr())
            {
                newSheet.worksheet.sheetPr = worksheet.sheetPr.Clone();
            }

            if(GetDefaultSheetView().pane != null)
            {
                CT_Pane oldPane = GetDefaultSheetView().pane;
                CT_Pane newPane = newSheet.GetPane();
                newPane.activePane = oldPane.activePane;
                newPane.state = oldPane.state;
                newPane.topLeftCell = oldPane.topLeftCell;
                newPane.xSplit = oldPane.xSplit;
                newPane.ySplit = oldPane.ySplit;
            }

            CopySheetImages(dest as XSSFWorkbook, newSheet);
        }

        public XSSFWorkbook GetWorkbook()
        {
            return (XSSFWorkbook) GetParent();
        }

        /// <summary>
        /// Create a pivot table using the AreaReference range on sourceSheet, 
        /// at the given position. If the source reference contains a sheet 
        /// name, it must match the sourceSheet
        /// </summary>
        /// <param name="source">location of pivot data</param>
        /// <param name="position">A reference to the top left cell where the 
        /// pivot table will start</param>
        /// <param name="sourceSheet">The sheet containing the source data, if 
        /// the source reference doesn't contain a sheet name</param>
        /// <returns>The pivot table</returns>
        /// <exception cref="ArgumentException">if source references a sheet 
        /// different than sourceSheet</exception>
        public XSSFPivotTable CreatePivotTable(AreaReference source, CellReference position, ISheet sourceSheet)
        {
            string sourceSheetName = source.FirstCell.SheetName;
            if(sourceSheetName != null
               && !sourceSheetName.Equals(sourceSheet.SheetName, StringComparison.InvariantCultureIgnoreCase))
            {
                throw new ArgumentException(
                    "The area is referenced in another sheet than the defined" +
                    " source sheet " + sourceSheet.SheetName + ".");
            }

            XSSFPivotTable.IPivotTableReferenceConfigurator refConfig = new PivotTableReferenceConfigurator1(source);
            return CreatePivotTable(position, sourceSheet, refConfig);
        }

        /// <summary>
        /// Create a pivot table using the AreaReference range, at the given 
        /// position. If the source reference contains a sheet name, that sheet
        /// is used, otherwise this sheet is assumed as the source sheet.
        /// </summary>
        /// <param name="source">location of pivot data</param>
        /// <param name="position">A reference to the top left cell where the 
        /// pivot table will start</param>
        /// <returns>The pivot table</returns>
        public XSSFPivotTable CreatePivotTable(AreaReference source, CellReference position)
        {
            string sourceSheetName = source.FirstCell.SheetName;
            if(sourceSheetName != null &&
               !sourceSheetName.Equals(SheetName, StringComparison.InvariantCultureIgnoreCase))
            {
                XSSFSheet sourceSheet = Workbook.GetSheet(sourceSheetName) as XSSFSheet;
                return CreatePivotTable(source, position, sourceSheet);
            }

            return CreatePivotTable(source, position, this);
        }

        /// <summary>
        /// Create a pivot table using the Name range reference on sourceSheet,
        /// at the given position. If the source reference contains a sheet 
        /// name, it must match the sourceSheet
        /// </summary>
        /// <param name="source">location of pivot data</param>
        /// <param name="position">A reference to the top left cell where the 
        /// pivot table will start</param>
        /// <param name="sourceSheet">The sheet containing the source data, 
        /// if the source reference doesn't contain a sheet name</param>
        /// <returns>The pivot table</returns>
        /// <exception cref="ArgumentException"></exception>
        public XSSFPivotTable CreatePivotTable(IName source, CellReference position, ISheet sourceSheet)
        {
            if(source.SheetName != null
               && !source.SheetName.Equals(sourceSheet.SheetName))
            {
                throw new ArgumentException(
                    "The named range references another sheet than the defined" +
                    " source sheet " + sourceSheet.SheetName + ".");
            }

            return CreatePivotTable(position, sourceSheet, new PivotTableReferenceConfigurator2(source));
        }

        /// <summary>
        /// Create a pivot table using the Name range, at the given position. 
        /// If the source reference contains a sheet name, that sheet is used, 
        /// otherwise this sheet is assumed as the source sheet.
        /// </summary>
        /// <param name="source">location of pivot data</param>
        /// <param name="position">A reference to the top left cell where the 
        /// pivot table will start</param>
        /// <returns>The pivot table</returns>
        public XSSFPivotTable CreatePivotTable(IName source, CellReference position)
        {
            return CreatePivotTable(
                source,
                position,
                GetWorkbook().GetSheet(source.SheetName));
        }

        /// <summary>
        /// Create a pivot table using the Table, at the given position. Tables
        /// are required to have a sheet reference, so no additional logic 
        /// around reference sheet is needed.
        /// </summary>
        /// <param name="source">location of pivot data</param>
        /// <param name="position">A reference to the top left cell where the 
        /// pivot table will start</param>
        /// <returns>The pivot table</returns>
        public XSSFPivotTable CreatePivotTable(ITable source, CellReference position)
        {
            return CreatePivotTable(
                position,
                GetWorkbook().GetSheet(source.SheetName),
                new PivotTableReferenceConfigurator3(source));
        }

        /// <summary>
        /// Returns all the pivot tables for this Sheet
        /// </summary>
        /// <returns></returns>
        public List<XSSFPivotTable> GetPivotTables()
        {
            List<XSSFPivotTable> tables = new List<XSSFPivotTable>();
            foreach(XSSFPivotTable table in GetWorkbook().PivotTables)
            {
                if(table.GetParent() == this)
                {
                    tables.Add(table);
                }
            }

            return tables;
        }

        public int GetColumnOutlineLevel(int columnIndex)
        {
            IColumn col = GetColumn(columnIndex);
            if(col == null)
            {
                return 0;
            }

            return col.OutlineLevel;
        }

        public bool IsDate1904()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Add ignored errors (usually to suppress them in the UI of a 
        /// consuming application).
        /// </summary>
        /// <param name="cell">Cell</param>
        /// <param name="ignoredErrorTypes">Types of error to ignore there.</param>
        public void AddIgnoredErrors(CellReference cell, params IgnoredErrorType[] ignoredErrorTypes)
        {
            AddIgnoredErrors(cell.FormatAsString(), ignoredErrorTypes);
        }

        /// <summary>
        /// Ignore errors across a range of cells.
        /// </summary>
        /// <param name="region">Range of cells.</param>
        /// <param name="ignoredErrorTypes">Types of error to ignore there.</param>
        public void AddIgnoredErrors(CellRangeAddress region, params IgnoredErrorType[] ignoredErrorTypes)
        {
            region.Validate(SpreadsheetVersion.EXCEL2007);
            AddIgnoredErrors(region.FormatAsString(), ignoredErrorTypes);
        }

        /// <summary>
        /// Returns the errors currently being ignored and the ranges where 
        /// they are ignored.
        /// </summary>
        /// <returns>Map of error type to the range(s) where they are ignored.</returns>
        public Dictionary<IgnoredErrorType, ISet<CellRangeAddress>> GetIgnoredErrors()
        {
            Dictionary<IgnoredErrorType, ISet<CellRangeAddress>> result =
                new Dictionary<IgnoredErrorType, ISet<CellRangeAddress>>();
            if(worksheet.IsSetIgnoredErrors())
            {
                foreach(CT_IgnoredError err in worksheet.ignoredErrors.ignoredError)
                {
                    foreach(IgnoredErrorType errType in GetErrorTypes(err))
                    {
                        if(!result.ContainsKey(errType))
                        {
                            result.Add(errType, new HashSet<CellRangeAddress>());
                        }

                        foreach(object ref1 in err.sqref)
                        {
                            result[errType].Add(CellRangeAddress.ValueOf(ref1.ToString()));
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Copies comment from one cell to another
        /// </summary>
        /// <param name="sourceCell">Cell with a comment to copy</param>
        /// <param name="targetCell">Cell to paste the comment to</param>
        /// <returns>Copied comment</returns>
        public IComment CopyComment(ICell sourceCell, ICell targetCell)
        {
            ValidateCellsForCopyComment(sourceCell, targetCell);

            XSSFComment sourceComment =
                sheetComments.FindCellComment(sourceCell.Address);
            CT_Comment sourceCtComment = sourceComment.GetCTComment();
            CT_Shape sourceCommentShape = GetVMLDrawing(false).FindCommentShape(
                sourceComment.Row,
                sourceComment.Column);

            CT_Comment targetCtComment = sheetComments.NewComment(targetCell.Address);
            targetCtComment.Set(sourceCtComment);

            CT_Shape targetCommentShape = GetVMLDrawing(false).newCommentShape();
            targetCommentShape.Set(sourceCommentShape);

            IComment targetComment = new XSSFComment(
                sheetComments,
                targetCtComment,
                targetCommentShape);

            targetCell.CellComment = targetComment;

            return targetComment;
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Read hyperlink relations, link them with CT_Hyperlink beans in this
        /// worksheet and Initialize the internal array of XSSFHyperlink objects
        /// </summary>
        /// <exception cref="POIXMLException"></exception>
        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        private void InitHyperlinks()
        {
            hyperlinks = new List<XSSFHyperlink>();

            if(!worksheet.IsSetHyperlinks())
            {
                return;
            }

            try
            {
                PackageRelationshipCollection hyperRels =
                    GetPackagePart().GetRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.Relation);

                // Turn each one into a XSSFHyperlink
                foreach(CT_Hyperlink hyperlink in worksheet.hyperlinks.hyperlink)
                {
                    PackageRelationship hyperRel = null;
                    if(hyperlink.id != null)
                    {
                        hyperRel = hyperRels.GetRelationshipByID(hyperlink.id);
                    }

                    hyperlinks.Add(new XSSFHyperlink(hyperlink, hyperRel));
                }
            }
            catch(InvalidFormatException e)
            {
                throw new POIXMLException(e);
            }
        }

        private void InitRows(CT_Worksheet worksheetParam)
        {
            _rows.Clear();
            tables = new Dictionary<string, XSSFTable>();
            sharedFormulas = new Dictionary<int, CT_CellFormula>();
            arrayFormulas = new List<CellRangeAddress>();
            if(0 < worksheetParam.sheetData.SizeOfRowArray())
            {
                foreach(CT_Row row in worksheetParam.sheetData.row)
                {
                    XSSFRow r = new XSSFRow(row, this);
                    if(!_rows.ContainsKey(r.RowNum))
                    {
                        _rows.Add(r.RowNum, r);
                    }
                }
            }
        }

        private void InitColumns(CT_Worksheet worksheetParam)
        {
            _columns.Clear();

            CT_Cols ctCols = worksheetParam.cols.FirstOrDefault() ?? worksheetParam.AddNewCols();

            if(ctCols.sizeOfColArray() > 0)
            {
                foreach(CT_Col column in ctCols.col)
                {
                    XSSFColumn c = new XSSFColumn(column, this);
                    if(!_columns.ContainsKey(c.ColumnNum))
                    {
                        _columns.Add(c.ColumnNum, c);
                    }
                }
            }
        }

        /// <summary>
        /// Create a new CT_Worksheet instance with all values set to defaults
        /// </summary>
        /// <returns>a new instance</returns>
        private static CT_Worksheet NewSheet()
        {
            CT_Worksheet worksheet = new CT_Worksheet();
            CT_SheetFormatPr ctFormat = worksheet.AddNewSheetFormatPr();
            ctFormat.defaultRowHeight = DEFAULT_ROW_HEIGHT;

            CT_SheetView ctView = worksheet.AddNewSheetViews().AddNewSheetView();
            ctView.workbookViewId = 0;

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

        /// <summary>
        /// Adds a merged region of cells (hence those cells form one).
        /// </summary>
        /// <param name="region">region (rowfrom/colfrom-rowto/colto) to merge</param>
        /// <param name="validate">whether to validate merged region</param>
        /// <returns>index of this region</returns>
        /// <exception cref="InvalidOperationException">if region intersects 
        /// with a multi-cell array formula or if region intersects with an 
        /// existing region on this sheet</exception>
        /// <exception cref="ArgumentException">if region contains fewer 
        /// than 2 cells</exception>
        private int AddMergedRegion(CellRangeAddress region, bool validate)
        {
            if(region.NumberOfCells < 2)
            {
                throw new ArgumentException("Merged region " +
                                            region.FormatAsString() + " must contain 2 or more cells");
            }

            region.Validate(SpreadsheetVersion.EXCEL2007);
            if(validate)
            {
                // throw InvalidOperationException if the argument
                // CellRangeAddress intersects with a multi-cell array formula
                // defined in this sheet
                ValidateArrayFormulas(region);
                // Throw InvalidOperationException if the argument
                // CellRangeAddress intersects with a merged region already
                // in this sheet 
                ValidateMergedRegions(region);
            }

            CT_MergeCells ctMergeCells = worksheet.IsSetMergeCells()
                ? worksheet.mergeCells
                : worksheet.AddNewMergeCells();
            CT_MergeCell ctMergeCell = ctMergeCells.AddNewMergeCell();
            ctMergeCell.@ref = region.FormatAsString();
            return ctMergeCells.sizeOfMergeCellArray() - 1;
        }

        /// <summary>
        /// Verify that the candidate region does not intersect with an 
        /// existing multi-cell array formula in this sheet
        /// </summary>
        /// <param name="region"></param>
        /// <exception cref="InvalidOperationException">if candidate region 
        /// intersects an existing array formula in this sheet</exception>
        private void ValidateArrayFormulas(CellRangeAddress region)
        {
            // FIXME: this may be faster if it looped over array formulas
            // directly rather than looping over each cell in the region and
            // searching if that cell belongs to an array formula
            int firstRow = region.FirstRow;
            int firstColumn = region.FirstColumn;
            int lastRow = region.LastRow;
            int lastColumn = region.LastColumn;
            // for each cell in sheet, if cell belongs to an array formula,
            // check if merged region intersects array formula cells
            for(int rowIn = firstRow; rowIn <= lastRow; rowIn++)
            {
                IRow row = GetRow(rowIn);
                if(row == null)
                {
                    continue;
                }

                for(int colIn = firstColumn; colIn <= lastColumn; colIn++)
                {
                    ICell cell = row.GetCell(colIn);
                    if(cell == null)
                    {
                        continue;
                    }

                    if(cell.IsPartOfArrayFormulaGroup)
                    {
                        CellRangeAddress arrayRange = cell.ArrayFormulaRange;
                        if(arrayRange.NumberOfCells > 1 && region.Intersects(arrayRange))
                        {
                            string msg = "The range " + region.FormatAsString() +
                                         " intersects with a multi-cell array formula. " +
                                         "You cannot merge cells of an array.";
                            throw new InvalidOperationException(msg);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Verify that none of the merged regions intersect a multi-cell array
        /// formula in this sheet
        /// </summary>
        private void CheckForMergedRegionsIntersectingArrayFormulas()
        {
            foreach(CellRangeAddress region in MergedRegions)
            {
                ValidateArrayFormulas(region);
            }
        }

        /// <summary>
        /// Verify that candidate region does not intersect with an existing 
        /// merged region in this sheet
        /// </summary>
        /// <param name="candidateRegion"></param>
        /// <exception cref="InvalidOperationException">if candidate region 
        /// intersects an existing merged region in this sheet</exception>
        private void ValidateMergedRegions(CellRangeAddress candidateRegion)
        {
            foreach(CellRangeAddress existingRegion in MergedRegions)
            {
                if(existingRegion.Intersects(candidateRegion))
                {
                    throw new InvalidOperationException(
                        "Cannot add merged region " + candidateRegion.FormatAsString() +
                        " to sheet because it overlaps with an existing merged " +
                        "region (" + existingRegion.FormatAsString() + ").");
                }
            }
        }

        /// <summary>
        /// Verify that no merged regions intersect another merged 
        /// region in this sheet.
        /// </summary>
        /// <exception cref="InvalidOperationException">if at least one region 
        /// intersects with another merged region in this sheet</exception>
        private void CheckForIntersectingMergedRegions()
        {
            List<CellRangeAddress> regions = MergedRegions;
            int size = regions.Count;
            for(int i = 0; i < size; i++)
            {
                CellRangeAddress region = regions[i];
                foreach(CellRangeAddress other in regions.Skip(i + 1)) //regions.subList(i+1, regions.size()
                {
                    if(region.Intersects(other))
                    {
                        string msg = "The range " + region.FormatAsString() +
                                     " intersects with another merged region " +
                                     other.FormatAsString() + " in this sheet";
                        throw new InvalidOperationException(msg);
                    }
                }
            }
        }

        private int GetLastKey(IList<int> keys)
        {
            _ = keys.Count;
            return keys[keys.Count - 1];
        }

        private int HeadMapCount(IList<int> keys, int rownum)
        {
            int count = 0;
            foreach(int key in keys)
            {
                if(key < rownum)
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

        private CT_SheetFormatPr GetSheetTypeSheetFormatPr()
        {
            return worksheet.IsSetSheetFormatPr() ? worksheet.sheetFormatPr : worksheet.AddNewSheetFormatPr();
        }

        private CT_SheetPr GetSheetTypeSheetPr()
        {
            if(worksheet.sheetPr == null)
            {
                worksheet.sheetPr = new CT_SheetPr();
            }

            return worksheet.sheetPr;
        }

        private CT_HeaderFooter GetSheetTypeHeaderFooter()
        {
            if(worksheet.headerFooter == null)
            {
                worksheet.headerFooter = new CT_HeaderFooter();
            }

            return worksheet.headerFooter;
        }

        /// <summary>
        /// returns all rows between startRow and endRow, inclusive. Rows 
        /// between startRow and endRow that haven't been created are not 
        /// included in result unless createRowIfMissing is true
        /// </summary>
        /// <param name="startRowNum">the first row number in this 
        /// sheet to return</param>
        /// <param name="endRowNum">the last row number in this 
        /// sheet to return</param>
        /// <param name="createRowIfMissing"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">if startRowNum and endRowNum 
        /// are not in ascending order</exception>
        private List<XSSFRow> GetRows(int startRowNum, int endRowNum, bool createRowIfMissing)
        {
            if(startRowNum > endRowNum)
            {
                throw new ArgumentException("getRows: startRowNum must be " +
                                            "less than or equal to endRowNum");
            }

            List<XSSFRow> rows = new List<XSSFRow>();
            if(createRowIfMissing)
            {
                for(int i = startRowNum; i <= endRowNum; i++)
                {
                    if(GetRow(i) is not XSSFRow row)
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

        /// <summary>
        /// Ensure CT_Worksheet.CT_SheetPr.CT_OutlinePr
        /// </summary>
        /// <returns></returns>
        private CT_OutlinePr EnsureOutlinePr()
        {
            CT_SheetPr sheetPr = worksheet.IsSetSheetPr()
                ? worksheet.sheetPr
                : worksheet.AddNewSheetPr();
            return sheetPr.IsSetOutlinePr()
                ? sheetPr.outlinePr
                : sheetPr.AddNewOutlinePr();

        }

        /// <summary>
        /// Do not leave the width attribute undefined (see #52186).
        /// </summary>
        /// <param name="ctCols"></param>
        private void SetColWidthAttribute(CT_Cols ctCols)
        {
            foreach(CT_Col col in ctCols.GetColList())
            {
                if(!col.IsSetWidth())
                {
                    col.width = DefaultColumnWidth;
                    col.customWidth = false;
                }
            }
        }

        private short GetMaxOutlineLevelRows()
        {
            short outlineLevel = 0;
            foreach(XSSFRow xrow in _rows.Values)
            {
                outlineLevel = xrow.GetCTRow().outlineLevel > outlineLevel
                    ? xrow.GetCTRow().outlineLevel
                    : outlineLevel;
            }

            return outlineLevel;
        }

        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        [Obsolete]
        private short GetMaxOutlineLevelCols()
        {
            CT_Cols ctCols = worksheet.GetColsArray(0);
            short outlineLevel = 0;
            foreach(CT_Col col in ctCols.GetColList())
            {
                outlineLevel = col.outlineLevel > outlineLevel
                    ? col.outlineLevel
                    : outlineLevel;
            }

            return outlineLevel;
        }

        private bool ListIsEmpty(List<CT_MergeCell> list)
        {
            foreach(CT_MergeCell mc in list)
            {
                if(mc != null)
                {
                    return false;
                }
            }

            return true;
        }

        private SortedDictionary<int, IColumn> GetAdjacentOutlineColumns(IColumn col)
        {
            SortedDictionary<int, IColumn> group = new SortedDictionary<int, IColumn>() { { col.ColumnNum, col } };

            for(int i = col.ColumnNum - 1; i >= 0; i--)
            {
                IColumn columnToTheLeft = GetColumn(i);

                if(columnToTheLeft == null || columnToTheLeft.OutlineLevel < col.OutlineLevel)
                {
                    break;
                }

                group.Add(i, columnToTheLeft);
            }

            int idx;

            for(idx = col.ColumnNum + 1; idx <= SpreadsheetVersion.EXCEL2007.LastColumnIndex; idx++)
            {
                IColumn columnToTheRight = GetColumn(idx);

                if(columnToTheRight == null || columnToTheRight.OutlineLevel < col.OutlineLevel)
                {
                    break;
                }

                group.Add(idx, columnToTheRight);
            }

            // add the column behind the group to set colapsed to
            group.Add(idx, GetColumn(idx, true));

            return group;
        }

        private CT_SheetView GetSheetTypeSheetView()
        {
            if(GetDefaultSheetView() == null)
            {
                GetSheetTypeSheetViews().SetSheetViewArray(0, new CT_SheetView());
            }

            return GetDefaultSheetView();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowIndex">the zero based row index to collapse</param>
        private void CollapseRow(int rowIndex)
        {
            XSSFRow row = (XSSFRow) GetRow(rowIndex);
            if(row != null)
            {
                int startRow = FindStartOfRowOutlineGroup(rowIndex);

                // Hide all the columns until the end of the group
                int lastRow = WriteHidden(row, startRow, true);
                if(GetRow(lastRow) != null)
                {
                    ((XSSFRow) GetRow(lastRow)).GetCTRow().collapsed = true;
                }
                else
                {
                    XSSFRow newRow = (XSSFRow) CreateRow(lastRow);
                    newRow.GetCTRow().collapsed = true;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowIndex">the zero based row index to find from</param>
        /// <returns></returns>
        private int FindStartOfRowOutlineGroup(int rowIndex)
        {
            // Find the start of the group.
            int level = ((XSSFRow) GetRow(rowIndex)).GetCTRow().outlineLevel;
            int currentRow = rowIndex;
            while(GetRow(currentRow) != null)
            {
                if(((XSSFRow) GetRow(currentRow)).GetCTRow().outlineLevel < level)
                {
                    return currentRow + 1;
                }

                currentRow--;
            }

            return currentRow;
        }

        private int WriteHidden(XSSFRow xRow, int rowIndex, bool hidden)
        {
            int level = xRow.GetCTRow().outlineLevel;
            for(IEnumerator it = GetRowEnumerator(); it.MoveNext();)
            {
                xRow = (XSSFRow) it.Current;
                if(xRow.GetCTRow().outlineLevel >= level)
                {
                    xRow.GetCTRow().hidden = hidden;
                    rowIndex++;
                }
            }

            return rowIndex;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowNumber">the zero based row index to expand</param>
        private void ExpandRow(int rowNumber)
        {
            if(rowNumber == -1)
            {
                return;
            }

            XSSFRow row = (XSSFRow) GetRow(rowNumber);
            // If it is already expanded do nothing.
            if(!row.GetCTRow().IsSetHidden())
            {
                return;
            }

            // Find the start of the group.
            int startIdx = FindStartOfRowOutlineGroup(rowNumber);

            // Find the end of the group.
            int endIdx = FindEndOfRowOutlineGroup(rowNumber);

            // expand: collapsed must be unset hidden bit Gets unset _if_
            // surrounding groups are expanded you can determine this by
            // looking at the hidden bit of the enclosing group. You will have
            // to look at the start and the end of the current group to
            // determine which is the enclosing group hidden bit only is
            // altered for this outline level. ie. don't un-collapse Contained
            // groups
            if(!IsRowGroupHiddenByParent(rowNumber))
            {
                for(int i = startIdx; i < endIdx; i++)
                {
                    if(row.GetCTRow().outlineLevel == ((XSSFRow) GetRow(i))
                       .GetCTRow()
                       .outlineLevel)
                    {
                        ((XSSFRow) GetRow(i)).GetCTRow().UnsetHidden();
                    }
                    else if(!IsRowGroupCollapsed(i))
                    {
                        ((XSSFRow) GetRow(i)).GetCTRow().UnsetHidden();
                    }
                }
            }

            // Write collapse field
            row = GetRow(endIdx) as XSSFRow;
            if(row != null)
            {
                CT_Row ctRow = row.GetCTRow();
                // This avoids an IndexOutOfBounds if multiple nested groups
                // are collapsed/expanded
                if(ctRow.collapsed)
                {
                    ctRow.UnsetCollapsed();
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row">the zero based row index to find from</param>
        /// <returns></returns>
        private bool IsRowGroupHiddenByParent(int row)
        {
            // Look out outline details of end
            int endLevel;
            bool endHidden;
            int endOfOutlineGroupIdx = FindEndOfRowOutlineGroup(row);
            if(GetRow(endOfOutlineGroupIdx) == null)
            {
                endLevel = 0;
                endHidden = false;
            }
            else
            {
                endLevel = ((XSSFRow) GetRow(endOfOutlineGroupIdx))
                    .GetCTRow().outlineLevel;
                endHidden = ((XSSFRow) GetRow(endOfOutlineGroupIdx))
                    .GetCTRow().hidden;
            }

            // Look out outline details of start
            int startLevel;
            bool startHidden;
            int startOfOutlineGroupIdx = FindStartOfRowOutlineGroup(row);
            if(startOfOutlineGroupIdx < 0
               || GetRow(startOfOutlineGroupIdx) == null)
            {
                startLevel = 0;
                startHidden = false;
            }
            else
            {
                startLevel = ((XSSFRow) GetRow(startOfOutlineGroupIdx))
                    .GetCTRow().outlineLevel;
                startHidden = ((XSSFRow) GetRow(startOfOutlineGroupIdx))
                    .GetCTRow().hidden;
            }

            if(endLevel > startLevel)
            {
                return endHidden;
            }

            return startHidden;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row">the zero based row index to find from</param>
        /// <returns></returns>
        private bool IsRowGroupCollapsed(int row)
        {
            int collapseRow = FindEndOfRowOutlineGroup(row) + 1;
            if(GetRow(collapseRow) == null)
            {
                return false;
            }

            return ((XSSFRow) GetRow(collapseRow)).GetCTRow().collapsed;
        }

        private void RemoveOverwrittenRows(int startRow, int endRow, int n)
        {
            XSSFVMLDrawing vml = GetVMLDrawing(false);
            List<int> rowsToRemove = new List<int>();
            List<CellAddress> commentsToRemove = new List<CellAddress>();
            List<CT_Row> ctRowsToRemove = new List<CT_Row>();
            // first remove all rows which will be overwritten
            foreach(KeyValuePair<int, XSSFRow> rowDict in _rows)
            {
                XSSFRow row = rowDict.Value;
                int rownum = row.RowNum;

                // check if we should remove this row as it will be overwritten
                // by the data later
                if(ShouldRemoveAtIndex(startRow, endRow, n, rownum))
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
                    if(sheetComments != null)
                    {
                        CT_CommentList lst = sheetComments.GetCTComments().commentList;
                        foreach(CT_Comment comment in lst.comment)
                        {
                            string strRef = comment.@ref;
                            CellAddress ref1 = new CellAddress(strRef);

                            // is this comment part of the current row?
                            if(ref1.Row == rownum)
                            {
                                //sheetComments.RemoveComment(strRef);
                                //vml.RemoveCommentShape(ref1.Row, ref1.Col);
                                commentsToRemove.Add(ref1);
                            }
                        }
                    }

                    foreach(CellAddress ref1 in commentsToRemove)
                    {
                        sheetComments.RemoveComment(ref1);
                        vml.RemoveCommentShape(ref1.Row, ref1.Column);
                    }

                    // FIXME: (performance optimization) this should be moved outside the for-loop so that hyperlinks only needs to be iterated over once.
                    // also remove any hyperlinks associated with this row
                    if(hyperlinks != null)
                    {
                        foreach(XSSFHyperlink link in new List<XSSFHyperlink>(hyperlinks))
                        {
                            CellReference ref1 = new CellReference(link.CellRef);
                            if(ref1.Row == rownum)
                            {
                                hyperlinks.Remove(link);
                            }
                        }
                    }
                }
            }

            foreach(int rowKey in rowsToRemove)
            {
                _rows.Remove(rowKey);
            }

            worksheet.sheetData.RemoveRows(ctRowsToRemove);
        }

        private void RemoveOverwrittenColumns(int startColumn, int endColumn, int n)
        {
            if(worksheet.cols.FirstOrDefault() is null)
            {
                throw new RuntimeException("There is no columns in XML part");
            }

            List<IColumn> columnsToRemove = new List<IColumn>();

            // first remove all columns which will be overwritten
            for(int i = startColumn + n; i <= endColumn + n; i++)
            {
                // check if we should remove this column as it will be overwritten
                // by the data later
                if(ShouldRemoveAtIndex(startColumn, endColumn, n, i))
                {
                    IColumn column = GetColumn(i, true);

                    columnsToRemove.Add(column);
                }
            }

            foreach(IColumn column in columnsToRemove)
            {
                RemoveColumn(column);
            }
        }

        private void RebuildRows()
        {
            //rebuild the _rows map
            Dictionary<int, XSSFRow> map = new Dictionary<int, XSSFRow>();
            foreach(XSSFRow r in _rows.Values)
            {
                map.Add(r.RowNum, r);
            }

            _rows.Clear();
            //_rows.putAll(map);
            foreach(KeyValuePair<int, XSSFRow> kv in map)
            {
                _rows.Add(kv.Key, kv.Value);
            }

            // Sort CTRows by index asc.
            // not found at poi 3.15
            if(worksheet.sheetData.row != null)
            {
                worksheet.sheetData.row.Sort((row1, row2) => row1.r.CompareTo(row2.r));
            }
        }

        private void RebuildColumns()
        {
            //rebuild the _columns map
            Dictionary<int, XSSFColumn> map = new Dictionary<int, XSSFColumn>();
            foreach(XSSFColumn c in _columns.Values)
            {
                map.Add(c.ColumnNum, c);
            }

            _columns.Clear();

            foreach(KeyValuePair<int, XSSFColumn> kv in map)
            {
                _columns.Add(kv.Key, kv.Value);
            }

            // Sort CT_Cols by index asc.
            worksheet.cols.FirstOrDefault()?.col.Sort((col1, col2) => col1.min.CompareTo(col2.min));
        }

        private void ShiftCommentsAndRows(int startRow, int endRow, int n, bool copyRowHeight)
        {
            SortedDictionary<XSSFComment, int> commentsToShift =
                new SortedDictionary<XSSFComment, int>(new ShiftCommentComparator(n));

            IEnumerable<CT_Shape> ctShapes = GetVMLDrawing(false)?.GetItems().OfType<CT_Shape>();

            foreach(KeyValuePair<int, XSSFRow> rowDict in _rows)
            {
                XSSFRow row = rowDict.Value;
                int rownum = row.RowNum;

                if(sheetComments != null)
                {
                    // calculate the new rownum
                    int newrownum = ShiftedRowOrColumnNumber(startRow, endRow, n, rownum);

                    // is there a change necessary for the current row?
                    if(newrownum != rownum)
                    {
                        List<CellAddress> commentAddresses = sheetComments.GetCellAddresses();
                        foreach(CellAddress cellAddress in commentAddresses)
                        {
                            if(cellAddress.Row == rownum)
                            {
                                XSSFComment oldComment = sheetComments
                                    .FindCellComment(cellAddress);
                                if(oldComment != null)
                                {
                                    var ctShape = oldComment.GetCTShape();

                                    if(ctShape == null && ctShapes != null)
                                    {
                                        ctShape = ctShapes.FirstOrDefault
                                        (x =>
                                            x.ClientData[0].row[0] == cellAddress.Row &&
                                            x.ClientData[0].column[0] == cellAddress.Column
                                        );
                                    }

                                    XSSFComment xssfComment =
                                        new XSSFComment(
                                            sheetComments,
                                            oldComment.GetCTComment(),
                                            ctShape);
                                    if(commentsToShift.ContainsKey(xssfComment))
                                    {
                                        commentsToShift[xssfComment] = newrownum;
                                    }
                                    else
                                    {
                                        commentsToShift.Add(xssfComment, newrownum);
                                    }
                                }
                            }
                        }
                    }
                }

                if(rownum < startRow || rownum > endRow)
                {
                    continue;
                }

                if(!copyRowHeight)
                {
                    row.Height = /*setter*/-1;
                }

                row.Shift(n);
            }

            // adjust all the affected comment-structures now
            // the Map is sorted and thus provides them in the order that we need here, 
            // i.e. from down to up if Shifting down, vice-versa otherwise
            foreach(KeyValuePair<XSSFComment, int> entry in commentsToShift)
            {
                entry.Key.Row = entry.Value;
            }

            RebuildRows();
        }

        private void ValidateCellsForCopyComment(ICell sourceCell, ICell targetCell)
        {
            if(sourceCell == null || targetCell == null)
            {
                throw new ArgumentException("Cells can not be null");
            }

            if(sourceCell.Sheet != targetCell.Sheet || sourceCell.Sheet != this)
            {
                throw new ArgumentException("Cells should belong to the same worksheet");
            }

            if(sheetComments == null)
            {
                throw new ArgumentException("Source cell doesn't have a comment");
            }

            if(sheetComments.FindCellComment(sourceCell.Address) == null)
            {
                throw new ArgumentException("Source cell doesn't have a comment");
            }
        }

        private void ShiftCommentsAndColumns(int startColumn, int endColumn, int n, bool copyColumnWidth)
        {
            SortedDictionary<XSSFComment, int> commentsToShift =
                new SortedDictionary<XSSFComment, int>(new ShiftCommentComparator(n));

            List<IColumn> columnsToDestroy = new List<IColumn>();

            for(int i = startColumn; i <= endColumn; i++)
            {
                if(!_columns.ContainsKey(i))
                {
                    columnsToDestroy.Add(CreateColumn(i));
                }
            }

            var sortedColumns = n < 0
                ? new SortedDictionary<int, XSSFColumn>(_columns)
                : new SortedDictionary<int, XSSFColumn>(_columns).Reverse();

            foreach(KeyValuePair<int, XSSFColumn> columnDict in sortedColumns)
            {
                XSSFColumn column = columnDict.Value;
                int columnNum = column.ColumnNum;

                if(sheetComments != null)
                {
                    // calculate the new columnNum
                    int newColumnNum = ShiftedRowOrColumnNumber(startColumn, endColumn, n, columnNum);

                    // is there a change necessary for the current column?
                    if(newColumnNum != columnNum)
                    {
                        List<CellAddress> commentAddresses = sheetComments.GetCellAddresses();
                        foreach(CellAddress cellAddress in commentAddresses)
                        {
                            if(cellAddress.Column == columnNum)
                            {
                                XSSFComment oldComment = sheetComments
                                    .FindCellComment(cellAddress);
                                if(oldComment != null)
                                {
                                    CT_Shape commentShape =
                                        GetVMLDrawing(false)
                                            .FindCommentShape(
                                                oldComment.Row,
                                                oldComment.Column);

                                    XSSFComment xssfComment =
                                        new XSSFComment(
                                            sheetComments,
                                            oldComment.GetCTComment(),
                                            commentShape);

                                    if(commentsToShift.ContainsKey(xssfComment))
                                    {
                                        commentsToShift[xssfComment] = newColumnNum;
                                    }
                                    else
                                    {
                                        commentsToShift.Add(xssfComment, newColumnNum);
                                    }
                                }
                            }
                        }
                    }
                }

                if(columnNum < startColumn || columnNum > endColumn)
                {
                    continue;
                }

                if(!copyColumnWidth)
                {
                    column.Width = -1;
                }

                column.Shift(n);
            }

            // adjust all the affected comment-structures now
            // the Map is sorted and thus provides them in the order that we need here, 
            // i.e. from right to left if Shifting right, vice-versa otherwise
            foreach(KeyValuePair<XSSFComment, int> entry in commentsToShift)
            {
                entry.Key.Column = entry.Value;
            }

            RebuildColumns();
            RebuildRowCells();
            DestroyColumns(columnsToDestroy);
        }

        private int GetIndexOfColumn(CT_Cols ctCols, CT_Col ctCol)
        {
            for(int i = 0; i < ctCols.sizeOfColArray(); i++)
            {
                if(ctCols.GetColArray(i).min == ctCol.min
                   && ctCols.GetColArray(i).max == ctCol.max)
                {
                    return i;
                }
            }

            return -1;
        }

        private void RebuildRowCells()
        {
            foreach(XSSFRow row in _rows.Values)
            {
                row.RebuildCells();
            }
        }

        private int ShiftedRowOrColumnNumber(int startIndex, int endIndex, int n, int index)
        {
            // no change if before any affected index
            if(index < startIndex && (n > 0 || (startIndex - index) > n))
            {
                return index;
            }

            // no change if After any affected index
            if(index > endIndex && (n < 0 || (index - endIndex) > n))
            {
                return index;
            }

            // index before and things are Moved up/left
            if(index < startIndex)
            {
                // index is Moved down/right by the Shifting
                return index + (endIndex - startIndex);
            }

            // index is After and things are Moved down/right
            if(index > endIndex)
            {
                // index is Moved up/left by the Shifting
                return index - (endIndex - startIndex);
            }

            // index is part of the Shifted block
            return index + n;
        }

        private CT_Selection GetSheetTypeSelection()
        {
            if(GetSheetTypeSheetView().SizeOfSelectionArray() == 0)
            {
                GetSheetTypeSheetView().InsertNewSelection(0);
            }

            return GetSheetTypeSheetView().GetSelectionArray(0);
        }

        /// <summary>
        /// Return the default sheet view. This is the last one if the sheet's 
        /// views, according to sec. 3.3.1.83 of the OOXML spec: "A single 
        /// sheet view defInition. When more than 1 sheet view is defined in 
        /// the file, it means that when opening the workbook, each sheet view 
        /// corresponds to a separate window within the spreadsheet 
        /// application, where each window is Showing the particular sheet. 
        /// Containing the same workbookViewId value, the last sheetView 
        /// defInition is loaded, and the others are discarded. When multiple 
        /// windows are viewing the same sheet, multiple sheetView elements 
        /// (with corresponding workbookView entries) are saved."
        /// </summary>
        /// <returns></returns>
        private CT_SheetView GetDefaultSheetView()
        {
            CT_SheetViews views = GetSheetTypeSheetViews();
            int sz = views == null ? 0 : views.sizeOfSheetViewArray();
            if(sz == 0)
            {
                return null;
            }

            return views.GetSheetViewArray(sz - 1);
        }

        private CT_PageSetUpPr GetSheetTypePageSetUpPr()
        {
            CT_SheetPr sheetPr = GetSheetTypeSheetPr();
            return sheetPr.IsSetPageSetUpPr() ? sheetPr.pageSetUpPr : sheetPr.AddNewPageSetUpPr();
        }

        private static bool ShouldRemoveAtIndex(int startIndex, int endIndex, int n, int currentIndex)
        {
            if(currentIndex >= (startIndex + n) && currentIndex <= (endIndex + n))
            {
                if(n > 0 && currentIndex > endIndex)
                {
                    return true;
                }
                else if(n < 0 && currentIndex < startIndex)
                {
                    return true;
                }
            }

            return false;
        }

        private CT_Pane GetPane()
        {
            if(GetDefaultSheetView().pane == null)
            {
                GetDefaultSheetView().AddNewPane();
            }

            return GetDefaultSheetView().pane;
        }

        private void SetSheetFormatPrOutlineLevelRow()
        {
            short maxLevelRow = GetMaxOutlineLevelRows();
            GetSheetTypeSheetFormatPr().outlineLevelRow = (byte) maxLevelRow;
        }

        private void SetSheetFormatPrOutlineLevelCol()
        {
            short maxLevelCol = GetMaxOutlineLevelCols();
            GetSheetTypeSheetFormatPr().outlineLevelCol = (byte) maxLevelCol;
        }

        private CT_SheetViews GetSheetTypeSheetViews()
        {
            if(worksheet.sheetViews == null)
            {
                worksheet.sheetViews = new CT_SheetViews();
                worksheet.sheetViews.AddNewSheetView();
            }

            return worksheet.sheetViews;
        }

        private CT_SheetProtection SafeGetProtectionField()
        {
            if(!IsSheetProtectionEnabled())
            {
                return worksheet.AddNewSheetProtection();
            }

            return worksheet.sheetProtection;
        }

        private bool IsSheetProtectionEnabled()
        {
            return worksheet.IsSetSheetProtection();
        }

        /// <summary>
        /// Also Creates cells if they don't exist
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        private ICellRange<ICell> GetCellRange(CellRangeAddress range)
        {
            int firstRow = range.FirstRow;
            int firstColumn = range.FirstColumn;
            int lastRow = range.LastRow;
            int lastColumn = range.LastColumn;
            int height = lastRow - firstRow + 1;
            int width = lastColumn - firstColumn + 1;
            List<ICell> temp = new List<ICell>(height * width);
            for(int rowIn = firstRow; rowIn <= lastRow; rowIn++)
            {
                for(int colIn = firstColumn; colIn <= lastColumn; colIn++)
                {
                    IRow row = GetRow(rowIn);
                    if(row == null)
                    {
                        row = CreateRow(rowIn);
                    }

                    ICell cell = row.GetCell(colIn);
                    if(cell == null)
                    {
                        cell = row.CreateCell(colIn);
                    }

                    temp.Add(cell);
                }
            }

            return SSCellRange<ICell>.Create(firstRow, firstColumn, height, width, temp, typeof(ICell));
        }

        private static string GetReferenceBuiltInRecord(
            string sheetName, int startC, int endC, int startR, int endR)
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

            string escapedName = SheetNameFormatter.Format(sheetName);

            string c = "";
            string r = "";

            if(startC != -1 || endC != -1)
            {
                string col1 = colRef.CellRefParts[2];
                string col2 = colRef2.CellRefParts[2];
                c = escapedName + "!$" + col1 + ":$" + col2;
            }

            if(startR != -1 || endR != -1)
            {
                string row1 = rowRef.CellRefParts[1];
                string row2 = rowRef2.CellRefParts[1];
                if(!row1.Equals("0") && !row2.Equals("0"))
                {
                    r = escapedName + "!$" + row1 + ":$" + row2;
                }
            }

            using var rng = ZString.CreateStringBuilder();
            rng.Append(c);
            if(rng.Length > 0 && r.Length > 0)
            {
                rng.Append(',');
            }

            rng.Append(r);
            return rng.ToString();
        }

        private CellRangeAddress GetRepeatingRowsOrColums(bool rows)
        {
            int sheetIndex = Workbook.GetSheetIndex(this);
            if(Workbook is not XSSFWorkbook xwb)
            {
                throw new RuntimeException("Workbook should not be null");
            }

            XSSFName name = xwb.GetBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
            if(name == null)
            {
                return null;
            }

            string refStr = name.RefersToFormula;
            if(refStr == null)
            {
                return null;
            }

            string[] parts = refStr.Split(",".ToCharArray());
            int maxRowIndex = SpreadsheetVersion.EXCEL2007.LastRowIndex;
            int maxColIndex = SpreadsheetVersion.EXCEL2007.LastColumnIndex;
            foreach(string part in parts)
            {
                CellRangeAddress range = CellRangeAddress.ValueOf(part);
                if((range.FirstColumn == 0
                    && range.LastColumn == maxColIndex)
                   || (range.FirstColumn == -1
                       && range.LastColumn == -1))
                {
                    if(rows)
                    {
                        return range;
                    }
                }
                else if((range.FirstRow == 0
                         && range.LastRow == maxRowIndex)
                        || (range.FirstRow == -1
                            && range.LastRow == -1))
                {
                    if(!rows)
                    {
                        return range;
                    }
                }
            }

            return null;
        }

        private void SetRepeatingRowsAndColumns(
            CellRangeAddress rowDef, CellRangeAddress colDef)
        {
            int col1 = -1;
            int col2 = -1;
            int row1 = -1;
            int row2 = -1;

            if(rowDef != null)
            {
                row1 = rowDef.FirstRow;
                row2 = rowDef.LastRow;
                if((row1 == -1 && row2 != -1)
                   || row1 < -1 || row2 < -1 || row1 > row2)
                {
                    throw new ArgumentException("Invalid row range specification");
                }
            }

            if(colDef != null)
            {
                col1 = colDef.FirstColumn;
                col2 = colDef.LastColumn;
                if((col1 == -1 && col2 != -1)
                   || col1 < -1 || col2 < -1 || col1 > col2)
                {
                    throw new ArgumentException(
                        "Invalid column range specification");
                }
            }

            int sheetIndex = Workbook.GetSheetIndex(this);

            bool removeAll = rowDef == null && colDef == null;
            if(Workbook is not XSSFWorkbook xwb)
            {
                throw new RuntimeException("Workbook should not be null");
            }

            XSSFName name = xwb.GetBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
            if(removeAll)
            {
                if(name != null)
                {
                    xwb.RemoveName(name);
                }

                return;
            }

            if(name == null)
            {
                name = xwb.CreateBuiltInName(
                    XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
            }

            string reference = GetReferenceBuiltInRecord(
                name.SheetName, col1, col2, row1, row2);
            name.RefersToFormula = reference;

            // If the print setup isn't currently defined, then add it
            //  in but without printer defaults
            // If it's already there, leave it as-is!
            if(worksheet.IsSetPageSetup() && worksheet.IsSetPageMargins())
            {
                // Everything we need is already there
            }
            else
            {
                // Have initial ones put in place
                PrintSetup.ValidSettings = false;
            }
        }

        private void CopySheetImages(XSSFWorkbook destWorkbook, XSSFSheet destSheet)
        {
            XSSFDrawing sheetDrawing = GetDrawingPatriarch();
            if(sheetDrawing != null)
            {
                IDrawing destDraw = destSheet.CreateDrawingPatriarch();
                IList<POIXMLDocumentPart> sheetPictures = sheetDrawing.GetRelations();
                Dictionary<string, uint> pictureIdMapping = new Dictionary<string, uint>();
                foreach(IEG_Anchor anchor in sheetDrawing.GetCTDrawing().CellAnchors)
                {
                    if(anchor is CT_TwoCellAnchor cellAnchor)
                    {
                        XSSFClientAnchor newAnchor = new XSSFClientAnchor(
                            (int) cellAnchor.from.colOff,
                            (int) cellAnchor.from.rowOff,
                            (int) cellAnchor.to.colOff,
                            (int) cellAnchor.to.rowOff,
                            cellAnchor.from.col,
                            cellAnchor.from.row,
                            cellAnchor.to.col,
                            cellAnchor.to.row);

                        if(cellAnchor.editAsSpecified)
                        {
                            switch(cellAnchor.editAs)
                            {
                                case ST_EditAs.twoCell:
                                    newAnchor.AnchorType = AnchorType.MoveAndResize;
                                    break;
                                case ST_EditAs.oneCell:
                                    newAnchor.AnchorType = AnchorType.MoveDontResize;
                                    break;
                                case ST_EditAs.absolute:
                                case ST_EditAs.NONE:
                                default:
                                    newAnchor.AnchorType = AnchorType.DontMoveAndResize;
                                    break;
                            }
                        }

                        string oldPictureId = anchor.picture?.blipFill?.blip.embed;
                        if(oldPictureId == null)
                        {
                            continue;
                        }

                        if(!pictureIdMapping.ContainsKey(oldPictureId))
                        {
                            XSSFPictureData srcPic = FindPicture(sheetPictures, oldPictureId);
                            if(srcPic != null && srcPic.PictureType != PictureType.None)
                            {
                                pictureIdMapping.Add(
                                    oldPictureId,
                                    (uint) destWorkbook.AddPicture(srcPic.Data, srcPic.PictureType));
                            }
                            else
                            {
                                continue; //Unable to find this picture, so skip it
                            }
                        }

                        destDraw.CreatePicture(newAnchor, (int) pictureIdMapping[oldPictureId]);
                    }
                }
            }
        }

        private XSSFPictureData FindPicture(IList<POIXMLDocumentPart> sheetPictures, string id)
        {
            foreach(POIXMLDocumentPart item in sheetPictures)
            {
                if(item.GetPackageRelationship().Id == id)
                {
                    return item as XSSFPictureData;
                }
            }

            return null;
        }

        private static void CopyRow(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow,
            IDictionary<int, ICellStyle> styleMap, bool keepFormulas, bool keepMergedRegion)
        {
            destRow.Height = srcRow.Height;
            if(!srcRow.GetCTRow().IsSetCustomHeight())
            {
                //Copying height sets the custom height flag, but Excel will
                //set a value for height even if it's auto-sized.
                destRow.GetCTRow().UnsetCustomHeight();
            }

            destRow.Hidden = srcRow.Hidden;
            destRow.Collapsed = srcRow.Collapsed;
            destRow.OutlineLevel = srcRow.OutlineLevel;

            if(srcRow.FirstCellNum < 0)
            {
                return; //Row has no cells, this sometimes happens with hidden or blank rows
            }

            for(int j = srcRow.FirstCellNum; j <= srcRow.LastCellNum; j++)
            {
                XSSFCell oldCell = (XSSFCell) srcRow.GetCell(j);
                XSSFCell newCell = (XSSFCell) destRow.GetCell(j);
                if(srcSheet.Workbook == destSheet.Workbook)
                {
                    newCell = (XSSFCell) destRow.GetCell(j);
                }

                if(oldCell != null)
                {
                    if(newCell == null)
                    {
                        newCell = (XSSFCell) destRow.CreateCell(j);
                    }

                    CopyCell(oldCell, newCell, styleMap, keepFormulas);

                    if(keepMergedRegion)
                    {
                        CellRangeAddress mergedRegion = srcSheet.GetMergedRegion(
                            new CellRangeAddress(srcRow.RowNum, srcRow.RowNum,
                                (short) oldCell.ColumnIndex,
                                (short) oldCell.ColumnIndex));

                        if(mergedRegion != null)
                        {
                            CellRangeAddress newMergedRegion = new CellRangeAddress(
                                mergedRegion.FirstRow,
                                mergedRegion.LastRow,
                                mergedRegion.FirstColumn,
                                mergedRegion.LastColumn);

                            if(!destSheet.IsMergedRegion(newMergedRegion))
                            {
                                destSheet.AddMergedRegion(newMergedRegion);
                            }
                        }
                    }
                }
            }
        }

        private static void CopyCell(ICell oldCell, ICell newCell, IDictionary<int, ICellStyle> styleMap,
            bool keepFormulas)
        {
            if(styleMap != null)
            {
                if(oldCell.CellStyle != null)
                {
                    if(oldCell.Sheet.Workbook == newCell.Sheet.Workbook)
                    {
                        newCell.CellStyle = oldCell.CellStyle;
                    }
                    else
                    {
                        int styleHashCode = oldCell.CellStyle.GetHashCode();
                        if (styleMap.TryGetValue(styleHashCode, out ICellStyle value))
                        {
                            newCell.CellStyle = value;
                        }
                        else
                        {
                            ICellStyle newCellStyle = newCell.Sheet.Workbook.CreateCellStyle();
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

            switch(oldCell.CellType)
            {
                case CellType.String:
                    XSSFRichTextString rts = oldCell.RichStringCellValue as XSSFRichTextString;
                    newCell.SetCellValue(rts);
                    if(rts != null)
                    {
                        for(int j = 0; j < rts.NumFormattingRuns; j++)
                        {
                            int startIndex = rts.GetIndexOfFormattingRun(j);
                            int endIndex;
                            if(j + 1 == rts.NumFormattingRuns)
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
                    if(keepFormulas)
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
                        catch
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

        /// <summary>
        /// Creates an empty XSSFPivotTable and Sets up all its relationships 
        /// including: pivotCacheDefInition, pivotCacheRecords
        /// </summary>
        /// <returns>a pivotTable</returns>
        private XSSFPivotTable CreatePivotTable()
        {
            XSSFWorkbook wb = GetWorkbook();
            List<XSSFPivotTable> pivotTables = wb.PivotTables;
            int tableId = GetWorkbook().PivotTables.Count + 1;
            //Create relationship between pivotTable and the worksheet
            XSSFPivotTable pivotTable = (XSSFPivotTable) CreateRelationship(
                XSSFRelation.PIVOT_TABLE,
                XSSFFactory.GetInstance(),
                tableId);
            pivotTable.SetParentSheet(this);
            pivotTables.Add(pivotTable);
            XSSFWorkbook workbook = GetWorkbook();

            //Create relationship between the pivot cache defintion and the workbook
            XSSFPivotCacheDefinition pivotCacheDefinition = (XSSFPivotCacheDefinition) workbook.CreateRelationship(
                XSSFRelation.PIVOT_CACHE_DEFINITION,
                XSSFFactory.GetInstance(),
                tableId);
            string rId = workbook.GetRelationId(pivotCacheDefinition);
            //Create relationship between pivotTable and pivotCacheDefInition
            //without creating a new instance
            PackagePart pivotPackagePart = pivotTable.GetPackagePart();
            pivotPackagePart.AddRelationship(
                pivotCacheDefinition.GetPackagePart().PartName,
                TargetMode.Internal,
                XSSFRelation.PIVOT_CACHE_DEFINITION.Relation);

            pivotTable.SetPivotCacheDefinition(pivotCacheDefinition);

            //Create pivotCache and Sets up it's relationship with the workbook
            pivotTable.SetPivotCache(new XSSFPivotCache(workbook.AddPivotCache(rId)));

            //Create relationship between pivotcacherecord and pivotcachedefInition
            XSSFPivotCacheRecords pivotCacheRecords = (XSSFPivotCacheRecords) pivotCacheDefinition.CreateRelationship(
                XSSFRelation.PIVOT_CACHE_RECORDS,
                XSSFFactory.GetInstance(),
                tableId);

            //Set relationships id for pivotCacheDefInition to pivotCacheRecords
            pivotTable.GetPivotCacheDefinition().GetCTPivotCacheDefinition().id =
                pivotCacheDefinition.GetRelationId(pivotCacheRecords);

            wb.PivotTables = /*setter*/pivotTables;

            return pivotTable;
        }

        /// <summary>
        /// Create a pivot table using the AreaReference or named/table range 
        /// on sourceSheet, at the given position. If the source reference 
        /// contains a sheet name, it must match the sourceSheet.
        /// </summary>
        /// <param name="position">A reference to the top left cell where the 
        /// pivot table will start</param>
        /// <param name="sourceSheet">The sheet containing the source data, 
        /// if the source reference doesn't contain a sheet name</param>
        /// <param name="refConfig"></param>
        /// <returns>The pivot table</returns>
        private XSSFPivotTable CreatePivotTable(CellReference position, ISheet sourceSheet,
            XSSFPivotTable.IPivotTableReferenceConfigurator refConfig)
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

        private void AddIgnoredErrors(string ref1, params IgnoredErrorType[] ignoredErrorTypes)
        {
            CT_IgnoredErrors ctIgnoredErrors = worksheet.IsSetIgnoredErrors()
                ? worksheet.ignoredErrors
                : worksheet.AddNewIgnoredErrors();
            CT_IgnoredError ctIgnoredError = ctIgnoredErrors.AddNewIgnoredError();
            XSSFIgnoredErrorHelper.AddIgnoredErrors(ctIgnoredError, ref1, ignoredErrorTypes);
        }

        private ISet<IgnoredErrorType> GetErrorTypes(CT_IgnoredError err)
        {
            ISet<IgnoredErrorType> result = new HashSet<IgnoredErrorType>();

            foreach(IgnoredErrorType errType in IgnoredErrorTypeValues.Values)
            {
                if(XSSFIgnoredErrorHelper.IsSet(errType, err))
                {
                    result.Add(errType);
                }
            }

            return result;
        }

        #endregion

        #region Helper classes

        public class PivotTableReferenceConfigurator2 : XSSFPivotTable.IPivotTableReferenceConfigurator
        {
            private readonly IName source;

            public PivotTableReferenceConfigurator2(IName source)
            {
                this.source = source;
            }

            public void ConfigureReference(CT_WorksheetSource wsSource)
            {
                wsSource.name = source.NameName;
            }
        }

        public class PivotTableReferenceConfigurator1 : XSSFPivotTable.IPivotTableReferenceConfigurator
        {
            private readonly AreaReference source;

            public PivotTableReferenceConfigurator1(AreaReference source)
            {
                this.source = source;
            }

            public void ConfigureReference(CT_WorksheetSource wsSource)
            {
                string[] firstCell = source.FirstCell.CellRefParts;
                string firstRow = firstCell[1];
                string firstCol = firstCell[2];
                string[] lastCell = source.LastCell.CellRefParts;
                string lastRow = lastCell[1];
                string lastCol = lastCell[2];
                string ref1 = firstCol + firstRow + ':' + lastCol + lastRow; //or just source.formatAsString()
                wsSource.@ref = ref1;
            }
        }

        public class PivotTableReferenceConfigurator3 : XSSFPivotTable.IPivotTableReferenceConfigurator
        {
            private readonly ITable source;

            public PivotTableReferenceConfigurator3(ITable source)
            {
                this.source = source;
            }

            public void ConfigureReference(CT_WorksheetSource wsSource)
            {
                wsSource.name = source.Name;
            }
        }

        private sealed class ShiftCommentComparator : IComparer<XSSFComment>
        {
            private readonly int shiftDir;

            public ShiftCommentComparator(int shiftDir)
            {
                this.shiftDir = shiftDir;
            }

            public int Compare(XSSFComment o1, XSSFComment o2)
            {
                int row1 = o1.Row;
                int row2 = o2.Row;

                if(row1 == row2)
                {
                    // ordering is not important when row is equal, but don't
                    // return zero to still get multiple comments per row into
                    // the map
                    return o1.GetHashCode() - o2.GetHashCode();
                }

                // when Shifting down, sort higher row-values first
                if(shiftDir > 0)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Width">In EMU</param>
        public void SetDefaultColWidth(int Width)
        {
            IFont ift = GetWorkbook().GetStylesSource().GetFontAt(0);
            Font ft = SheetUtil.IFont2Font(ift);
            var rt = TextMeasurer.MeasureSize("0", new TextOptions(ft));
            double MDW = rt.Width + 1; //MaximumDigitWidth

            worksheet.sheetFormatPr.defaultColWidth = Width / Units.EMU_PER_PIXEL / MDW;
            worksheet.cols.Clear();
        }

        [Obsolete("")]
        public double MaximumDigitWidth
        {
            set;
            get;
        } = 7.0;

        [Obsolete("")]
        public double GetDefaultColWidthInPixel()
        {
            double fontwidth;
            double width_px;

            var width = worksheet.sheetFormatPr.defaultColWidth; //string length with padding
            if(width != 0.0)
            {
                double widthInPx = width * MaximumDigitWidth;
                width_px = widthInPx + (8 - widthInPx % 8); // round up to the nearest multiple of 8 pixels
            }
            else if(worksheet.sheetFormatPr.baseColWidth != 0)
            {
                double MDW = MaximumDigitWidth;
                var length = worksheet.sheetFormatPr.baseColWidth; //string length with out padding
                fontwidth = Math.Truncate((length * MDW + 5) / MDW * 256) / 256;
                double tmp = 256 * fontwidth + Math.Truncate(128 / MDW);
                width_px = Math.Truncate((tmp / 256) * MDW) + 3; // +3 ???
            }
            else
            {
                double widthInPx = DefaultColumnWidth * MaximumDigitWidth;
                width_px = widthInPx + (8 - widthInPx % 8); // round up to the nearest multiple of 8 pixels
            }

            return width_px;
        }

        [Obsolete("")]
        public XSSFClientAnchor CreateClientAnchor(
            int dx1
            , int dy1
            , int dx2
            , int dy2
        )
        {
            int left = Math.Min(dx1, dx2);
            int right = Math.Max(dx1, dx2);
            int top = Math.Min(dy1, dy2);
            int bottom = Math.Max(dy1, dy2);

            CT_Marker mk1 = EMUtoMarker(left, top);
            CT_Marker mk2 = EMUtoMarker(right, bottom);

            if(mk1.colOff >= 0 && mk1.rowOff >= 0 && mk2.colOff >= 0 && mk2.rowOff >= 0)
            {
                return new XSSFClientAnchor(mk1, mk2, left, top, right, bottom);
            }

            return null;
        }

        [Obsolete("")]
        private CT_Marker EMUtoMarker(int x, int y)
        {
            CT_Marker mkr = new CT_Marker();
            mkr.colOff = EMUtoColOff(x, out int col);
            mkr.col = col;
            mkr.rowOff = EMUtoRowOff(y, out int row);
            mkr.row = row;
            return mkr;
        }

        [Obsolete("")]
        public int EMUtoRowOff(
            int y
            , out int cell
        )
        {
            double height;
            cell = 0;
            for(int iRow = 0; iRow < SpreadsheetVersion.EXCEL2007.MaxRows; iRow++)
            {
                height = _rows.TryGetValue(iRow, out XSSFRow row) ? row.HeightInPoints
                                                 : DefaultRowHeightInPoints;
                if (y >= Units.ToEMU(height))
                {
                    y -= Units.ToEMU(height);
                    cell++;
                }
                else
                {
                    return y;
                }
            }

            return -1;
        }

        [Obsolete("")]
        public int EMUtoColOff(
            int x
            , out int cell
        )
        {

            double width_px;
            cell = 0;
            for(int iCol = 0; iCol < SpreadsheetVersion.EXCEL2007.MaxColumns; iCol++)
            {
                width_px = GetDefaultColWidthInPixel();
                foreach(var cols in worksheet.cols)
                {
                    foreach(var col in cols.col)
                    {
                        if(col.min <= iCol + 1 && iCol + 1 <= col.max)
                        {
                            width_px = col.width * MaximumDigitWidth;
                            goto lblforbreak;
                        }
                    }
                }

                lblforbreak:
                int EMUwidth = Units.PixelToEMU((int) Math.Round(width_px, 1));
                if(x >= EMUwidth)
                {
                    x -= EMUwidth;
                    cell++;
                }
                else
                {
                    return x;
                }
            }

            return -1;
        }

        #endregion

        public CellRangeAddressList GetCells(string cellranges)
        {
            return CellRangeAddressList.Parse(cellranges);
        }
    }
}
