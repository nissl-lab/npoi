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
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFSheet : ISheet
    {
        /*package*/
        private XSSFSheet _sh;
        private SXSSFWorkbook _workbook;
        //private TreeMap<Integer, SXSSFRow> _rows = new TreeMap<Integer, SXSSFRow>();
        //private SheetDataWriter _writer;
        private int _randomAccessWindowSize = SXSSFWorkbook.DEFAULT_WINDOW_SIZE;
        private object _autoSizeColumnTracker;
        private int outlineLevelRow = 0;
        private int lastFlushedRowNumber = -1;
        private bool allFlushed = false;


        public SXSSFSheet(SXSSFWorkbook workbook, XSSFSheet xSheet)
        {
            _workbook = workbook;
            _sh = xSheet;
            _writer = workbook.createSheetDataWriter();
            SetRandomAccessWindowSize(_workbook.GetRandomAccessWindowSize());
            _autoSizeColumnTracker = new AutoSizeColumnTracker(this);
    }
        public void SetRandomAccessWindowSize(int value)
        {
            if (value == 0 || value < -1)
            {
                throw new ArgumentException("RandomAccessWindowSize must be either -1 or a positive integer");
            }
            _randomAccessWindowSize = value;
        }


        public bool Autobreaks
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public int[] ColumnBreaks
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public int DefaultColumnWidth
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public short DefaultRowHeight
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public float DefaultRowHeightInPoints
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool DisplayFormulas
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool DisplayGridlines
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool DisplayGuts
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool DisplayRowColHeadings
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool DisplayZeros
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public IDrawing DrawingPatriarch
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public int FirstRowNum
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool FitToPage
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public IFooter Footer
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool ForceFormulaRecalculation
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public IHeader Header
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool HorizontallyCenter
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool IsActive
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool IsPrintGridlines
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool IsRightToLeft
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool IsSelected
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public int LastRowNum
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public short LeftCol
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public int NumMergedRegions
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public PaneInformation PaneInformation
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public int PhysicalNumberOfRows
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public IPrintSetup PrintSetup
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool Protect
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public CellRangeAddress RepeatingColumns
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public CellRangeAddress RepeatingRows
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public int[] RowBreaks
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public bool RowSumsBelow
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool RowSumsRight
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool ScenarioProtect
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public ISheetConditionalFormatting SheetConditionalFormatting
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public string SheetName
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public short TabColorIndex
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public short TopRow
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool VerticallyCenter
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public IWorkbook Workbook
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public int AddMergedRegion(CellRangeAddress region)
        {
            throw new NotImplementedException();
        }

        public void AddValidationData(IDataValidation dataValidation)
        {
            throw new NotImplementedException();
        }

        public void AutoSizeColumn(int column)
        {
            throw new NotImplementedException();
        }

        public void AutoSizeColumn(int column, bool useMergedCells)
        {
            throw new NotImplementedException();
        }

        public IRow CopyRow(int sourceIndex, int targetIndex)
        {
            throw new NotImplementedException();
        }

        public ISheet CopySheet(string Name)
        {
            throw new NotImplementedException();
        }

        public ISheet CopySheet(string Name, bool copyStyle)
        {
            throw new NotImplementedException();
        }

        public IDrawing CreateDrawingPatriarch()
        {
            throw new NotImplementedException();
        }

        public void CreateFreezePane(int colSplit, int rowSplit)
        {
            throw new NotImplementedException();
        }

        public void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
        {
            throw new NotImplementedException();
        }

        public IRow CreateRow(int rownum)
        {
            throw new NotImplementedException();
        }

        public void CreateSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PanePosition activePane)
        {
            throw new NotImplementedException();
        }

        public IComment GetCellComment(int row, int column)
        {
            throw new NotImplementedException();
        }

        public int GetColumnOutlineLevel(int columnIndex)
        {
            throw new NotImplementedException();
        }

        public ICellStyle GetColumnStyle(int column)
        {
            throw new NotImplementedException();
        }

        public int GetColumnWidth(int columnIndex)
        {
            throw new NotImplementedException();
        }

        public float GetColumnWidthInPixels(int columnIndex)
        {
            throw new NotImplementedException();
        }

        public IDataValidationHelper GetDataValidationHelper()
        {
            throw new NotImplementedException();
        }

        public List<IDataValidation> GetDataValidations()
        {
            throw new NotImplementedException();
        }

        public IEnumerator GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public double GetMargin(MarginType margin)
        {
            throw new NotImplementedException();
        }

        public CellRangeAddress GetMergedRegion(int index)
        {
            throw new NotImplementedException();
        }

        public IRow GetRow(int rownum)
        {
            throw new NotImplementedException();
        }

        public IEnumerator GetRowEnumerator()
        {
            throw new NotImplementedException();
        }

        public void GroupColumn(int fromColumn, int toColumn)
        {
            throw new NotImplementedException();
        }

        public void GroupRow(int fromRow, int toRow)
        {
            throw new NotImplementedException();
        }

        public bool IsColumnBroken(int column)
        {
            throw new NotImplementedException();
        }

        public bool IsColumnHidden(int columnIndex)
        {
            throw new NotImplementedException();
        }

        public bool IsMergedRegion(CellRangeAddress mergedRegion)
        {
            throw new NotImplementedException();
        }

        public bool IsRowBroken(int row)
        {
            throw new NotImplementedException();
        }

        public void ProtectSheet(string password)
        {
            throw new NotImplementedException();
        }

        public ICellRange<ICell> RemoveArrayFormula(ICell cell)
        {
            throw new NotImplementedException();
        }

        public void RemoveColumnBreak(int column)
        {
            throw new NotImplementedException();
        }

        public void RemoveMergedRegion(int index)
        {
            throw new NotImplementedException();
        }

        public void RemoveRow(IRow row)
        {
            throw new NotImplementedException();
        }

        public void RemoveRowBreak(int row)
        {
            throw new NotImplementedException();
        }

        public void SetActive(bool value)
        {
            throw new NotImplementedException();
        }

        public void SetActiveCell(int row, int column)
        {
            throw new NotImplementedException();
        }

        public void SetActiveCellRange(List<CellRangeAddress8Bit> cellranges, int activeRange, int activeRow, int activeColumn)
        {
            throw new NotImplementedException();
        }

        public void SetActiveCellRange(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            throw new NotImplementedException();
        }

        public ICellRange<ICell> SetArrayFormula(string formula, CellRangeAddress range)
        {
            throw new NotImplementedException();
        }

        public IAutoFilter SetAutoFilter(CellRangeAddress range)
        {
            throw new NotImplementedException();
        }

        public void SetColumnBreak(int column)
        {
            throw new NotImplementedException();
        }

        public void SetColumnGroupCollapsed(int columnNumber, bool collapsed)
        {
            throw new NotImplementedException();
        }

        public void SetColumnHidden(int columnIndex, bool hidden)
        {
            throw new NotImplementedException();
        }

        public void SetColumnWidth(int columnIndex, int width)
        {
            throw new NotImplementedException();
        }

        public void SetDefaultColumnStyle(int column, ICellStyle style)
        {
            throw new NotImplementedException();
        }

        public void SetMargin(MarginType margin, double size)
        {
            throw new NotImplementedException();
        }

        public void SetRowBreak(int row)
        {
            throw new NotImplementedException();
        }

        public void SetRowGroupCollapsed(int row, bool collapse)
        {
            throw new NotImplementedException();
        }

        public void SetZoom(int numerator, int denominator)
        {
            throw new NotImplementedException();
        }

        public void ShiftRows(int startRow, int endRow, int n)
        {
            throw new NotImplementedException();
        }

        public void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight)
        {
            throw new NotImplementedException();
        }

        public void ShowInPane(short toprow, short leftcol)
        {
            throw new NotImplementedException();
        }

        public void ShowInPane(int toprow, int leftcol)
        {
            throw new NotImplementedException();
        }

        public void UngroupColumn(int fromColumn, int toColumn)
        {
            throw new NotImplementedException();
        }

        public void UngroupRow(int fromRow, int toRow)
        {
            throw new NotImplementedException();
        }

        public bool IsDate1904()
        {
            throw new NotImplementedException();
        }
    }
}
