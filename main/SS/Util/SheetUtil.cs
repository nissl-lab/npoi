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
    using System.Drawing;
    using System.Windows.Forms;
    using System.Collections.Generic;

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
                if (cellRangeAddress.FirstRow == sourceRow.RowNum)
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
            else
            {
                newRow = sheet.CreateRow(targetRowIndex);
            }

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
                if (cellRangeAddress.FirstRow == sourceRow.RowNum)
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


        /**
         * Compute width of a single cell
         *
         * @param cell the cell whose width is to be calculated
         * @param defaultCharWidth the width of a single character
         * @param formatter formatter used to prepare the text to be measured
         * @param useMergedCells    whether to use merged cells
         * @return  the width in pixels
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

            //AttributedString str;
            //TextLayout layout;

            double width = -1;
            using (Bitmap bmp = new Bitmap(1,1))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                if (cellType == CellType.String)
                {
                    IRichTextString rt = cell.RichStringCellValue;
                    String[] lines = rt.String.Split("\n".ToCharArray());
                    for (int i = 0; i < lines.Length; i++)
                    {
                        String txt = lines[i] + defaultChar;

                        //str = new AttributedString(txt);
                        //copyAttributes(font, str, 0, txt.length());
                        windowsFont = IFont2Font(font);
                        if (rt.NumFormattingRuns > 0)
                        {
                            // TODO: support rich text fragments
                        }

                        //layout = new TextLayout(str.getIterator(), fontRenderContext);
                        if (style.Rotation != 0)
                        {
                            /*
                             * Transform the text using a scale so that it's height is increased by a multiple of the leading,
                             * and then rotate the text before computing the bounds. The scale results in some whitespace around
                             * the unrotated top and bottom of the text that normally wouldn't be present if unscaled, but
                             * is added by the standard Excel autosize.
                             */
                            //AffineTransform trans = new AffineTransform();
                            //trans.concatenate(AffineTransform.getRotateInstance(style.Rotation*2.0*Math.PI/360.0));
                            //trans.concatenate(
                            //    AffineTransform.getScaleInstance(1, fontHeightMultiple)
                            //    );
                            double angle = style.Rotation * 2.0 * Math.PI / 360.0;
                            SizeF sf = g.MeasureString(txt, windowsFont);
                            double x1 = Math.Abs(sf.Height * Math.Sin(angle));
                            double x2 = Math.Abs(sf.Width * Math.Cos(angle));
                            double w = Math.Round(x1 + x2, 0, MidpointRounding.ToEven);
                            width = Math.Max(width, (w / colspan / defaultCharWidth) * 2 + cell.CellStyle.Indention);
                            //width = Math.Max(width,
                            //                 ((layout.getOutline(trans).getBounds().getWidth()/colspan)/defaultCharWidth) +
                            //                 cell.getCellStyle().getIndention());
                        }
                        else
                        {
                            //width = Math.Max(width,
                            //                 ((layout.getBounds().getWidth()/colspan)/defaultCharWidth) +
                            //                 cell.getCellStyle().getIndention());
                            double w = Math.Round(g.MeasureString(txt, windowsFont).Width, 0, MidpointRounding.ToEven);
                            width = Math.Max(width, (w / colspan / defaultCharWidth) * 2 + cell.CellStyle.Indention);
                        }
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
                        //layout = new TextLayout(str.getIterator(), fontRenderContext);
                        if (style.Rotation != 0)
                        {
                            /*
                             * Transform the text using a scale so that it's height is increased by a multiple of the leading,
                             * and then rotate the text before computing the bounds. The scale results in some whitespace around
                             * the unrotated top and bottom of the text that normally wouldn't be present if unscaled, but
                             * is added by the standard Excel autosize.
                             */
                            //AffineTransform trans = new AffineTransform();
                            //trans.concatenate(AffineTransform.getRotateInstance(style.getRotation()*2.0*Math.PI/360.0));
                            //trans.concatenate(
                            //    AffineTransform.getScaleInstance(1, fontHeightMultiple)
                            //    );
                            //width = Math.max(width,
                            //                 ((layout.getOutline(trans).getBounds().getWidth()/colspan)/defaultCharWidth) +
                            //                 cell.getCellStyle().getIndention());
                            double angle = style.Rotation * 2.0 * Math.PI / 360.0;
                            SizeF sf = g.MeasureString(txt, windowsFont);
                            double x1 = sf.Height * Math.Sin(angle);
                            double x2 = sf.Width * Math.Cos(angle);
                            double w = Math.Round(x1 + x2, 0, MidpointRounding.ToEven);
                            width = Math.Max(width, (w / colspan / defaultCharWidth) * 2 + cell.CellStyle.Indention);
                        }
                        else
                        {
                            //width = Math.max(width,
                            //                 ((layout.getBounds().getWidth()/colspan)/defaultCharWidth) +
                            //                 cell.getCellStyle().getIndention());
                            double w = Math.Round(g.MeasureString(txt, windowsFont).Width, 0, MidpointRounding.ToEven);
                            width = Math.Max(width, (w * 1.0 / colspan / defaultCharWidth) * 2 + cell.CellStyle.Indention);
                        }
                    }
                }
            }
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
         * @return  the width in pixels
         */

        public static double GetColumnWidth(ISheet sheet, int column, bool useMergedCells)
        {
            //AttributedString str;
            //TextLayout layout;

            IWorkbook wb = sheet.Workbook;
            DataFormatter formatter = new DataFormatter();
            IFont defaultFont = wb.GetFontAt((short) 0);

            //str = new AttributedString((defaultChar));
            //copyAttributes(defaultFont, str, 0, 1);
            //layout = new TextLayout(str.Iterator, fontRenderContext);
            //int defaultCharWidth = (int)layout.Advance;
            Font font = IFont2Font(defaultFont);
            int defaultCharWidth = TextRenderer.MeasureText("" + new String(defaultChar, 1), font).Width;
            //DummyEvaluator dummyEvaluator = new DummyEvaluator();

            double width = -1;
            foreach (IRow row in sheet)
            {
                ICell cell = row.GetCell(column);

                if (cell == null)
                {
                    continue;
                }

                double cellWidth = GetCellWidth(cell, defaultCharWidth, formatter, useMergedCells);
                width = Math.Max(width, cellWidth);
            }
            return width;
        }


        /**
         * Compute width of a column based on a subset of the rows and return the result
         *
         * @param sheet the sheet to calculate
         * @param column    0-based index of the column
         * @param useMergedCells    whether to use merged cells
         * @param firstRow  0-based index of the first row to consider (inclusive)
         * @param lastRow   0-based index of the last row to consider (inclusive)
         * @return  the width in pixels
         */
        public static double GetColumnWidth(ISheet sheet, int column, bool useMergedCells, int firstRow, int lastRow)
        {
            IWorkbook wb = sheet.Workbook;
            DataFormatter formatter = new DataFormatter();

            IFont defaultFont = wb.GetFontAt((short)0);

            //str = new AttributedString((defaultChar));
            //copyAttributes(defaultFont, str, 0, 1);
            //layout = new TextLayout(str.Iterator, fontRenderContext);
            //int defaultCharWidth = (int)layout.Advance;
            Font font = IFont2Font(defaultFont);
            int defaultCharWidth = TextRenderer.MeasureText("" + new String(defaultChar, 1), font).Width;

            double width = -1;
            for (int rowIdx = firstRow; rowIdx <= lastRow; ++rowIdx)
            {
                IRow row = sheet.GetRow(rowIdx);
                if (row != null)
                {

                    ICell cell = row.GetCell(column);

                    if (cell == null)
                    {
                        continue;
                    }

                    double cellWidth = GetCellWidth(cell, defaultCharWidth, formatter, useMergedCells);
                    width = Math.Max(width, cellWidth);
                }
            }
            return width;
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
        //    if (font.Underline == (byte)FontUnderlineType.SINGLE) str.AddAttribute(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON, startIdx, endIdx);           
        //}

        /// <summary>
        /// Convert HSSFFont to Font.
        /// </summary>
        /// <param name="font1">The font.</param>
        /// <returns></returns>
        internal static Font IFont2Font(IFont font1)
        {
            FontStyle style = FontStyle.Regular;
            if (font1.Boldweight == (short)FontBoldWeight.Bold)
            {
                style |= FontStyle.Bold;
            }
            if (font1.IsItalic)
                style |= FontStyle.Italic;
            if (font1.Underline == FontUnderlineType.Single)
            {
                style |= FontStyle.Underline;
            }
            Font font = new Font(font1.FontName, font1.FontHeightInPoints, style, GraphicsUnit.Point);
            return font;
            //return new System.Drawing.Font(font1.FontName, font1.FontHeightInPoints);
        }
        public static bool ContainsCell(CellRangeAddress cr, int rowIx, int colIx)
        {
            if (cr.FirstRow <= rowIx && cr.LastRow >= rowIx
                    && cr.FirstColumn <= colIx && cr.LastColumn >= colIx)
            {
                return true;
            }
            return false;
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
