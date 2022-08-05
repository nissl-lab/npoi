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
    using System.Collections.Generic;
    using SixLabors.Fonts;
    using System.Linq;

    /**
     * Helper methods for when working with Usermodel sheets
     *
     * @author Yegor Kozlov
     */
    public class SheetUtil
    {

        // /**
        // * Excel measures columns in units of 1/256th of a character width
        // * but the docs say nothing about what particular character is used.
        // * '0' looks to be a good choice.
        // */
        private static char defaultChar = '0';

        // /**
        // * This is the multiple that the font height is scaled by when determining the
        // * boundary of rotated text.
        // */
        //private static double fontHeightMultiple = 2.0;

        /**
         *  Dummy formula Evaluator that does nothing.
         *  YK: The only reason of having this class is that
         *  {@link NPOI.SS.UserModel.DataFormatter#formatCellValue(NPOI.SS.UserModel.Cell)}
         *  returns formula string for formula cells. Dummy Evaluator Makes it to format the cached formula result.
         *
         *  See Bugzilla #50021 
         */
        private static IFormulaEvaluator dummyEvaluator = new DummyEvaluator();
        public class DummyEvaluator : IFormulaEvaluator
        {
            public void ClearAllCachedResultValues() { }
            public void NotifySetFormula(ICell cell) { }
            public void NotifyDeleteCell(ICell cell) { }
            public void NotifyUpdateCell(ICell cell) { }
            public CellValue Evaluate(ICell cell) { return null; }
            public ICell EvaluateInCell(ICell cell) { return null; }
            public bool IgnoreMissingWorkbooks { get; set; }
            public void SetupReferencedWorkbooks(Dictionary<String, IFormulaEvaluator> workbooks) { }
            public void EvaluateAll() { }

            public CellType EvaluateFormulaCell(ICell cell)
            {
                return cell.CachedFormulaResultType;
            }

            public bool DebugEvaluationOutputForNextEval
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
        }
        public static IRow CopyRow(ISheet sourceSheet, int sourceRowIndex, ISheet targetSheet, int targetRowIndex)
        {
            // Get the source / new row
            IRow newRow = targetSheet.GetRow(targetRowIndex);
            IRow sourceRow = sourceSheet.GetRow(sourceRowIndex);

            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null)
            {
                targetSheet.RemoveRow(newRow);
            }
            newRow = targetSheet.CreateRow(targetRowIndex);
            if (sourceRow == null)
                throw new ArgumentNullException("source row doesn't exist");
            // Loop through source columns to add to new row
            for (int i = sourceRow.FirstCellNum; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                ICell oldCell = sourceRow.GetCell(i);

                // If the old cell is null jump to next cell
                if (oldCell == null)
                {
                    continue;
                }
                ICell newCell = newRow.CreateCell(i);

                if (oldCell.CellStyle != null)
                {
                    // apply style from old cell to new cell 
                    newCell.CellStyle = oldCell.CellStyle;
                }

                // If there is a cell comment, copy
                if (oldCell.CellComment != null)
                {
                    newCell.CellComment = oldCell.CellComment;
                }

                // If there is a cell hyperlink, copy
                if (oldCell.Hyperlink != null)
                {
                    newCell.Hyperlink = oldCell.Hyperlink;
                }

                // Set the cell data type
                newCell.SetCellType(oldCell.CellType);

                // Set the cell data value
                switch (oldCell.CellType)
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

            // If there are are any merged regions in the source row, copy to new row
            for (int i = 0; i < sourceSheet.NumMergedRegions; i++)
            {
                CellRangeAddress cellRangeAddress = sourceSheet.GetMergedRegion(i);

                if (cellRangeAddress != null && cellRangeAddress.FirstRow == sourceRow.RowNum)
                {
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
                            (newRow.RowNum +
                                    (cellRangeAddress.LastRow - cellRangeAddress.FirstRow
                                            )),
                            cellRangeAddress.FirstColumn,
                            cellRangeAddress.LastColumn);
                    targetSheet.AddMergedRegion(newCellRangeAddress);
                }
            }
            return newRow;
        }
        public static IRow CopyRow(ISheet sheet, int sourceRowIndex, int targetRowIndex)
        {
            if (sourceRowIndex == targetRowIndex)
                throw new ArgumentException("sourceIndex and targetIndex cannot be same");
            // Get the source / new row
            IRow newRow = sheet.GetRow(targetRowIndex);
            IRow sourceRow = sheet.GetRow(sourceRowIndex);

            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null)
            {
                sheet.ShiftRows(targetRowIndex, sheet.LastRowNum, 1);
            }
            newRow = sheet.CreateRow(targetRowIndex);
            newRow.Height = sourceRow.Height;   //copy row height

            // Loop through source columns to add to new row
            for (int i = sourceRow.FirstCellNum; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                ICell oldCell = sourceRow.GetCell(i);

                // If the old cell is null jump to next cell
                if (oldCell == null)
                {
                    continue;
                }
                ICell newCell = newRow.CreateCell(i);

                if (oldCell.CellStyle != null)
                {
                    // apply style from old cell to new cell 
                    newCell.CellStyle = oldCell.CellStyle;
                }

                // If there is a cell comment, copy
                if (oldCell.CellComment != null)
                {
                    newCell.CellComment = oldCell.CellComment;
                }

                // If there is a cell hyperlink, copy
                if (oldCell.Hyperlink != null)
                {
                    newCell.Hyperlink = oldCell.Hyperlink;
                }

                // Set the cell data type
                newCell.SetCellType(oldCell.CellType);

                // Set the cell data value
                switch (oldCell.CellType)
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

            // If there are are any merged regions in the source row, copy to new row
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress cellRangeAddress = sheet.GetMergedRegion(i);
                if (cellRangeAddress != null && cellRangeAddress.FirstRow == sourceRow.RowNum)
                {
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
                            (newRow.RowNum +
                                    (cellRangeAddress.LastRow - cellRangeAddress.FirstRow
                                            )),
                            cellRangeAddress.FirstColumn,
                            cellRangeAddress.LastColumn);
                    sheet.AddMergedRegion(newCellRangeAddress);
                }
            }
            return newRow;
        }

        public static double GetRowHeight(IRow row, bool useMergedCells, int firstColumnIdx, int lastColumnIdx)
        {
            double width = -1;

            for (int cellIdx = firstColumnIdx; cellIdx <= lastColumnIdx; ++cellIdx)
            {
                ICell cell = row.GetCell(cellIdx);
                if (row != null && cell != null)
                {
                    double cellWidth = GetCellHeight(cell, useMergedCells);
                    width = Math.Max(width, cellWidth);
                }
            }

            return width;
        }

        public static double GetRowHeight(ISheet sheet, int rowIdx, bool useMergedCells, int firstColumnIdx, int lastColumnIdx)
        {
            IRow row = sheet.GetRow(rowIdx);
            return GetRowHeight(row, useMergedCells, firstColumnIdx, lastColumnIdx);
        }

        public static double GetRowHeight(IRow row, bool useMergedCells)
        {
            if (row == null)
            {
                return -1;
            }

            double rowHeight = -1;

            foreach (var cell in row.Cells)
            {
                double cellHeight = GetCellHeight(cell, useMergedCells);
                rowHeight = Math.Max(rowHeight, cellHeight);
            }

            return rowHeight;
        }

        public static double GetRowHeight(ISheet sheet, int rowIdx, bool useMergedCells)
        {
            IRow row = sheet.GetRow(rowIdx);

            return GetRowHeight(row, useMergedCells);
        }

        public static double GetCellHeight(ICell cell, bool useMergedCells)
        {
            ICell cellToMeasure = useMergedCells ? GetFirstCellFromMergedRegion(cell) : cell;

            double stringHeight = GetActualHeight(cellToMeasure);
            int numberOfRowsInMergedRegion = useMergedCells ? GetNumberOfRowsInMergedRegion(cellToMeasure) : 1;

            return GetCellConetntHeight(stringHeight, numberOfRowsInMergedRegion);
        }

        private static ICell GetFirstCellFromMergedRegion(ICell cell)
        {
            foreach (var region in cell.Sheet.MergedRegions)
            {
                if (region.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return cell.Sheet.GetRow(region.FirstRow).GetCell(region.FirstColumn);
                }
            }

            return cell;
        }

        private static double GetActualHeight(ICell cell)
        {
            string stringValue = GetCellStringValue(cell);
            Font windowsFont = GetWindowsFont(cell);

            if (cell.CellStyle.Rotation != 0)
            {
                return GetRotatedContentHeight(cell, stringValue, windowsFont);
            }

            return GetContentHeight(stringValue, windowsFont);
        }

        private static int GetNumberOfRowsInMergedRegion(ICell cell)
        {
            foreach (var region in cell.Sheet.MergedRegions)
            {
                if (region.IsInRange(cell.RowIndex, cell.ColumnIndex))
                {
                    return 1 + region.LastColumn - region.FirstColumn;
                }
            }

            return 1;
        }

        private static double GetCellConetntHeight(double actualHeight, int numberOfRowsInMergedRegion)
        {
            return Math.Max(-1, actualHeight / numberOfRowsInMergedRegion);
        }

        private static string GetCellStringValue(ICell cell)
        {
            CellType cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;

            if (cellType == CellType.String)
            {
                return cell.RichStringCellValue.String;
            }

            if (cellType == CellType.Boolean)
            {
                return cell.BooleanCellValue.ToString().ToUpper() + defaultChar;
            }

            if (cellType == CellType.Numeric)
            {
                string stringValue;

                try
                {
                    DataFormatter formatter = new DataFormatter();
                    stringValue = formatter.FormatCellValue(cell, dummyEvaluator);
                }
                catch (Exception)
                {
                    stringValue = cell.NumericCellValue.ToString();
                }

                return stringValue + defaultChar;
            }

            return null;
        }

        private static Font GetWindowsFont(ICell cell)
        {
            var wb = cell.Sheet.Workbook;
            var style = cell.CellStyle;
            var font = wb.GetFontAt(style.FontIndex);

            return IFont2Font(font);
        }

        private static double GetRotatedContentHeight(ICell cell, string stringValue, Font windowsFont)
        {
            var angle = cell.CellStyle.Rotation * 2.0 * Math.PI / 360.0;
            var measureResult = TextMeasurer.Measure(stringValue, new TextOptions(windowsFont));

            var x1 = Math.Abs(measureResult.Height * Math.Cos(angle));
            var x2 = Math.Abs(measureResult.Width * Math.Sin(angle));

            return Math.Round(x1 + x2, 0, MidpointRounding.ToEven);
        }

        private static double GetContentHeight(string stringValue, Font windowsFont)
        {
            var measureResult = TextMeasurer.Measure(stringValue, new TextOptions(windowsFont));
            
            return Math.Round(measureResult.Height, 0, MidpointRounding.ToEven);
        }

        /**
         * Compute width of a single cell
         *
         * @param cell the cell whose width is to be calculated
         * @param defaultCharWidth the width of a single character
         * @param formatter formatter used to prepare the text to be measured
         * @param useMergedCells    whether to use merged cells
         * @return  the width in pixels or -1 if cell is empty
         */
        public static double GetCellWidth(ICell cell, int defaultCharWidth, DataFormatter formatter, bool useMergedCells)
        {
            ISheet sheet = cell.Sheet;
            IWorkbook wb = sheet.Workbook;
            IRow row = cell.Row;
            int column = cell.ColumnIndex;

            int colspan = 1;
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress region = sheet.GetMergedRegion(i);
                if (ContainsCell(region, row.RowNum, column))
                {
                    if (!useMergedCells)
                    {
                        // If we're not using merged cells, skip this one and move on to the next.
                        return -1;
                    }
                    cell = row.GetCell(region.FirstColumn);
                    colspan = 1 + region.LastColumn - region.FirstColumn;
                }
            }

            ICellStyle style = cell.CellStyle;
            CellType cellType = cell.CellType;
            IFont defaultFont = wb.GetFontAt((short)0);
            Font windowsFont = IFont2Font(defaultFont);
            // for formula cells we compute the cell width for the cached formula result
            if (cellType == CellType.Formula) cellType = cell.CachedFormulaResultType;

            IFont font = wb.GetFontAt(style.FontIndex);

            double width = -1;
            {
                if (cellType == CellType.String)
                {
                    IRichTextString rt = cell.RichStringCellValue;
                    String[] lines = rt.String.Split("\n".ToCharArray());
                    for (int i = 0; i < lines.Length; i++)
                    {
                        String txt = lines[i] + defaultChar;

                        //AttributedString str = new AttributedString(txt);
                        //copyAttributes(font, str, 0, txt.length());
                        windowsFont = IFont2Font(font);
                        if (rt.NumFormattingRuns > 0)
                        {
                            // TODO: support rich text fragments
                        }

                        width = GetCellWidth(defaultCharWidth, colspan, style, width, txt, windowsFont, cell);
                    }
                }
                else
                {
                    String sval = null;
                    if (cellType == CellType.Numeric)
                    {
                        // Try to get it formatted to look the same as excel
                        try
                        {
                            sval = formatter.FormatCellValue(cell, dummyEvaluator);
                        }
                        catch (Exception)
                        {
                            sval = cell.NumericCellValue.ToString();
                        }
                    }
                    else if (cellType == CellType.Boolean)
                    {
                        sval = cell.BooleanCellValue.ToString().ToUpper();
                    }
                    if (sval != null)
                    {
                        String txt = sval + defaultChar;
                        //str = new AttributedString(txt);
                        //copyAttributes(font, str, 0, txt.length());
                        windowsFont = IFont2Font(font);
                        width = GetCellWidth(defaultCharWidth, colspan, style, width, txt, windowsFont, cell);
                    }
                }
            }
            return width;
        }

        private static double GetCellWidth(int defaultCharWidth, int colspan,
            ICellStyle style, double width, string str, Font windowsFont, ICell cell)
        {
            //Rectangle bounds;
            double actualWidth;
            FontRectangle sf = TextMeasurer.Measure(str, new TextOptions(windowsFont));
            if (style.Rotation != 0)
            {
                double angle = style.Rotation * 2.0 * Math.PI / 360.0;
                double x1 = Math.Abs(sf.Height * Math.Sin(angle));
                double x2 = Math.Abs(sf.Width * Math.Cos(angle));
                actualWidth = Math.Round(x1 + x2, 0, MidpointRounding.ToEven);
                //bounds = layout.getOutline(trans).getBounds();
            }
            else
            {
                //bounds = layout.getBounds();
                actualWidth = Math.Round(sf.Width, 0, MidpointRounding.ToEven);                
            }
            // entireWidth accounts for leading spaces which is excluded from bounds.getWidth()
            //double frameWidth = bounds.getX() + bounds.getWidth();
            //width = Math.max(width, ((frameWidth / colspan) / defaultCharWidth) + style.getIndention());
            width = Math.Max(width, (actualWidth / colspan / defaultCharWidth) + cell.CellStyle.Indention);
            return width;
        }

        // /**
        // * Drawing context to measure text
        // */
        //private static FontRenderContext fontRenderContext = new FontRenderContext(null, true, true);

        /**
         * Compute width of a column and return the result
         *
         * @param sheet the sheet to calculate
         * @param column    0-based index of the column
         * @param useMergedCells    whether to use merged cells
         * @return  the width in pixels or -1 if all cells are empty
         */
        public static double GetColumnWidth(ISheet sheet, int column, bool useMergedCells)
        {
            return GetColumnWidth(sheet, column, useMergedCells, sheet.FirstRowNum, sheet.LastRowNum);
        }
        /**
         * Compute width of a column based on a subset of the rows and return the result
         *
         * @param sheet the sheet to calculate
         * @param column    0-based index of the column
         * @param useMergedCells    whether to use merged cells
         * @param firstRow  0-based index of the first row to consider (inclusive)
         * @param lastRow   0-based index of the last row to consider (inclusive)
         * @param maxRows   limit the scope to maxRows rows to speed up the function, or leave 0 (optional)
         * @return  the width in pixels or -1 if cell is empty
         */
        public static double GetColumnWidth(ISheet sheet, int column, bool useMergedCells, int firstRow, int lastRow, int maxRows = 0)
        {
            DataFormatter formatter = new DataFormatter();
            int defaultCharWidth = GetDefaultCharWidth(sheet.Workbook);

            // No need to explore the whole sheet: explore only the first maxRows lines
            if (maxRows > 0 && lastRow - firstRow > maxRows) lastRow = firstRow + maxRows;

            double width = -1;
            for (int rowIdx = firstRow; rowIdx <= lastRow; ++rowIdx)
            {
                IRow row = sheet.GetRow(rowIdx);
                if (row != null)
                {
                    double cellWidth = GetColumnWidthForRow(row, column, defaultCharWidth, formatter, useMergedCells);
                    width = Math.Max(width, cellWidth);
                }
            }
            return width;
        }

        /**
         * Get default character width
         *
         * @param wb the workbook to get the default character width from
         * @return default character width
         */
        public static int GetDefaultCharWidth(IWorkbook wb)
        {
            IFont defaultFont = wb.GetFontAt((short)0);

            //AttributedString str = new AttributedString(String.valueOf(defaultChar));
            //copyAttributes(defaultFont, str, 0, 1);
            //TextLayout layout = new TextLayout(str.getIterator(), fontRenderContext);
            //int defaultCharWidth = (int)layout.getAdvance();
            Font font = IFont2Font(defaultFont);
            return (int)Math.Ceiling(TextMeasurer.Measure(new string(defaultChar, 1), new TextOptions(font)).Width);
        }

        /**
         * Compute width of a single cell in a row
         * Convenience method for {@link getCellWidth}
         *
         * @param row the row that contains the cell of interest
         * @param column the column number of the cell whose width is to be calculated
         * @param defaultCharWidth the width of a single character
         * @param formatter formatter used to prepare the text to be measured
         * @param useMergedCells    whether to use merged cells
         * @return  the width in pixels or -1 if cell is empty
         */
        private static double GetColumnWidthForRow(
                IRow row, int column, int defaultCharWidth, DataFormatter formatter, bool useMergedCells)
        {
            if (row == null)
            {
                return -1;
            }

            ICell cell = row.GetCell(column);

            if (cell == null)
            {
                return -1;
            }

            return GetCellWidth(cell, defaultCharWidth, formatter, useMergedCells);
        }

        /**
         * Check if the Fonts are installed correctly so that Java can compute the size of
         * columns. 
         * 
         * If a Cell uses a Font which is not available on the operating system then Java may 
         * fail to return useful Font metrics and thus lead to an auto-computed size of 0.
         * 
         *  This method allows to check if computing the sizes for a given Font will succeed or not.
         *
         * @param font The Font that is used in the Cell
         * @return true if computing the size for this Font will succeed, false otherwise
         */
        public static bool CanComputeColumnWidth(IFont font)
        {
            //AttributedString str = new AttributedString("1w");
            //copyAttributes(font, str, 0, "1w".length());

            //TextLayout layout = new TextLayout(str.getIterator(), fontRenderContext);
            //if (layout.getBounds().getWidth() > 0)
            //{
            //    return true;
            //}

            return true;
        }
        // /**
        // * Copy text attributes from the supplied Font to Java2D AttributedString
        // */
        //private static void copyAttributes(IFont font, AttributedString str, int startIdx, int endIdx)
        //{
        //    str.AddAttribute(TextAttribute.FAMILY, font.FontName, startIdx, endIdx);
        //    str.AddAttribute(TextAttribute.SIZE, (float)font.FontHeightInPoints);
        //    if (font.Boldweight == (short)FontBoldWeight.BOLD) str.AddAttribute(TextAttribute.WEIGHT, TextAttribute.WEIGHT_BOLD, startIdx, endIdx);
        //    if (font.IsItalic) str.AddAttribute(TextAttribute.POSTURE, TextAttribute.POSTURE_OBLIQUE, startIdx, endIdx);
        //    TODO-Fonts: not supported: if (font.Underline == (byte)FontUnderlineType.SINGLE) str.AddAttribute(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON, startIdx, endIdx);           
        //}

        /// <summary>
        /// Convert HSSFFont to Font.
        /// </summary>
        /// <param name="font1">The font.</param>
        /// <returns></returns>
        internal static Font IFont2Font(IFont font1)
        {
            FontStyle style = FontStyle.Regular;
            if (font1.IsBold)
            {
                style |= FontStyle.Bold;
            }
            if (font1.IsItalic)
                style |= FontStyle.Italic;

            /* TODO-Fonts: not supported
            if (font1.Underline == FontUnderlineType.Single)
            {
                style |= FontStyle.Underline;
            }
            */

            // Try to find font in system fonts. If we can not find out,
            // use "Arial". TODO-Fonts: More fallbacks.
            SixLabors.Fonts.FontFamily fontFamily;

            if (false == SystemFonts.TryGet(font1.FontName, out fontFamily))
            {
                if (false == SystemFonts.TryGet("Arial", out fontFamily))
                {
                    fontFamily = SystemFonts.Families.First();
                }
            }
            
            Font font = new Font(fontFamily, (float)font1.FontHeightInPoints, style);
            return font;
        }
        /// <summary>
        /// Check if the cell is in the specified cell range
        /// </summary>
        /// <param name="cr">the cell range to check in</param>
        /// <param name="rowIx">the row to check</param>
        /// <param name="colIx">the column to check</param>
        /// <returns>return true if the range contains the cell [rowIx, colIx]</returns>
        [Obsolete("deprecated 3.15 beta 2. Use {@link CellRangeAddressBase#isInRange(int, int)}.")]
        public static bool ContainsCell(CellRangeAddress cr, int rowIx, int colIx)
        {
            return cr.IsInRange(rowIx, colIx);
        }

        /**
         * Generate a valid sheet name based on the existing one. Used when cloning sheets.
         *
         * @param srcName the original sheet name to
         * @return clone sheet name
         */
        public static String GetUniqueSheetName(IWorkbook wb, String srcName)
        {
            if (wb.GetSheetIndex(srcName) == -1)
            {
                return srcName;
            }
            int uniqueIndex = 2;
            String baseName = srcName;
            int bracketPos = srcName.LastIndexOf('(');
            if (bracketPos > 0 && srcName.EndsWith(")"))
            {
                String suffix = srcName.Substring(bracketPos + 1, srcName.Length - bracketPos - 2);
                try
                {
                    uniqueIndex = Int32.Parse(suffix.Trim());
                    uniqueIndex++;
                    baseName = srcName.Substring(0, bracketPos).Trim();
                }
                catch (FormatException)
                {
                    // contents of brackets not numeric
                }
            }
            while (true)
            {
                // Try and find the next sheet name that is unique
                String index = (uniqueIndex++).ToString();
                String name;
                if (baseName.Length + index.Length + 2 < 31)
                {
                    name = baseName + " (" + index + ")";
                }
                else
                {
                    name = baseName.Substring(0, 31 - index.Length - 2) + "(" + index + ")";
                }

                //If the sheet name is unique, then Set it otherwise Move on to the next number.
                if (wb.GetSheetIndex(name) == -1)
                {
                    return name;
                }
            }
        }

        /**
         * Return the cell, taking account of merged regions. Allows you to find the
         *  cell who's contents are Shown in a given position in the sheet.
         * 
         * <p>If the cell at the given co-ordinates is a merged cell, this will
         *  return the primary (top-left) most cell of the merged region.</p>
         * <p>If the cell at the given co-ordinates is not in a merged region,
         *  then will return the cell itself.</p>
         * <p>If there is no cell defined at the given co-ordinates, will return
         *  null.</p>
         */
        public static ICell GetCellWithMerges(ISheet sheet, int rowIx, int colIx)
        {
            IRow r = sheet.GetRow(rowIx);
            if (r != null)
            {
                ICell c = r.GetCell(colIx);
                if (c != null)
                {
                    // Normal, non-merged cell
                    return c;
                }
            }

            for (int mr = 0; mr < sheet.NumMergedRegions; mr++)
            {
                CellRangeAddress mergedRegion = sheet.GetMergedRegion(mr);
                if (mergedRegion.IsInRange(rowIx, colIx))
                {
                    // The cell wanted is in this merged range
                    // Return the primary (top-left) cell for the range
                    r = sheet.GetRow(mergedRegion.FirstRow);
                    if (r != null)
                    {
                        return r.GetCell(mergedRegion.FirstColumn);
                    }
                }
            }

            // If we Get here, then the cell isn't defined, and doesn't
            //  live within any merged regions
            return null;
        }

    }
}
