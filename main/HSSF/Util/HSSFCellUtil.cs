/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.HSSF.Util
{
    using System;
    using System.Collections;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;
    using NPOI.HSSF.Record;
    using NPOI.SS.Util;

    /// <summary>
    /// Various utility functions that make working with a cells and rows easier.  The various
    /// methods that deal with style's allow you to Create your HSSFCellStyles as you need them.
    /// When you apply a style change to a cell, the code will attempt to see if a style already
    /// exists that meets your needs.  If not, then it will Create a new style.  This is to prevent
    /// creating too many styles.  there is an upper limit in Excel on the number of styles that
    /// can be supported.
    /// @author     Eric Pugh epugh@upstate.com
    /// </summary>
    [Obsolete("deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil} instead.")]
    public class HSSFCellUtil
    {
        private HSSFCellUtil()
        {
            // no instances of this class
        }

        /// <summary>
        /// Get a row from the spreadsheet, and Create it if it doesn't exist.
        /// </summary>
        /// <param name="rowCounter">The 0 based row number</param>
        /// <param name="sheet">The sheet that the row is part of.</param>
        /// <returns>The row indicated by the rowCounter</returns>
        [Obsolete("@deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil#getRow} instead.")]
        public static IRow GetRow(int rowIndex, HSSFSheet sheet)
        {
            return (HSSFRow)CellUtil.GetRow(rowIndex, sheet);
        }


        /// <summary>
        /// Get a specific cell from a row. If the cell doesn't exist,
        /// </summary>
        /// <param name="row">The row that the cell is part of</param>
        /// <param name="column">The column index that the cell is in.</param>
        /// <returns>The cell indicated by the column.</returns>
        [Obsolete("@deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil#getCell} instead.")]
        public static ICell GetCell(IRow row, int columnIndex)
        {
            return (HSSFCell)CellUtil.GetCell(row, columnIndex);
        }


        /// <summary>
        /// Creates a cell, gives it a value, and applies a style if provided
        /// </summary>
        /// <param name="row">the row to Create the cell in</param>
        /// <param name="column">the column index to Create the cell in</param>
        /// <param name="value">The value of the cell</param>
        /// <param name="style">If the style is not null, then Set</param>
        /// <returns>A new HSSFCell</returns>
        [Obsolete("@deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil#createCell} instead.")]
        public static ICell CreateCell(IRow row, int column, String value, HSSFCellStyle style)
        {
            return (HSSFCell)CellUtil.CreateCell(row, column, value, style);
        }


        /// <summary>
        /// Create a cell, and give it a value.
        /// </summary>
        /// <param name="row">the row to Create the cell in</param>
        /// <param name="column">the column index to Create the cell in</param>
        /// <param name="value">The value of the cell</param>
        /// <returns>A new HSSFCell.</returns>
        [Obsolete("@deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil#createCell} instead.")]
        public static ICell CreateCell(IRow row, int column, String value)
        {
            return CreateCell(row, column, value, null);
        }

        /// <summary>
        /// Take a cell, and align it.
        /// </summary>
        /// <param name="cell">the cell to Set the alignment for</param>
        /// <param name="workbook">The workbook that is being worked with.</param>
        /// <param name="align">the column alignment to use.</param>
        [Obsolete("@deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil#setAlignment} instead.")]
        public static void SetAlignment(ICell cell, HSSFWorkbook workbook, short align)
        {
            CellUtil.SetAlignment(cell, (HorizontalAlignment)align);
        }

        /// <summary>
        /// Take a cell, and apply a font to it
        /// </summary>
        /// <param name="cell">the cell to Set the alignment for</param>
        /// <param name="workbook">The workbook that is being worked with.</param>
        /// <param name="font">The HSSFFont that you want to Set...</param>
        [Obsolete("@deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil#setFont} instead.")]
        public static void SetFont(ICell cell, HSSFWorkbook workbook, HSSFFont font)
        {
            CellUtil.SetFont(cell, font);
        }

        /**
         *  This method attempt to find an already existing HSSFCellStyle that matches
         *  what you want the style to be. If it does not find the style, then it
         *  Creates a new one. If it does Create a new one, then it applies the
         *  propertyName and propertyValue to the style. This is necessary because
         *  Excel has an upper limit on the number of Styles that it supports.
         *
         *@param  workbook               The workbook that is being worked with.
         *@param  propertyName           The name of the property that is to be
         *      changed.
         *@param  propertyValue          The value of the property that is to be
         *      changed.
         *@param  cell                   The cell that needs it's style changes
         *@exception  NestableException  Thrown if an error happens.
         */
        [Obsolete("@deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil#setCellStyleProperty} instead.")]
        public static void SetCellStyleProperty(ICell cell, HSSFWorkbook workbook, String propertyName, Object propertyValue)
        {
            CellUtil.SetCellStyleProperty(cell, propertyName, propertyValue);
        }

        /// <summary>
        /// Looks for text in the cell that should be unicode, like alpha; and provides the
        /// unicode version of it.
        /// </summary>
        /// <param name="cell">The cell to check for unicode values</param>
        /// <returns>transalted to unicode</returns>
        [Obsolete("@deprecated 3.15 beta2. Removed in 3.17. Use {@link org.apache.poi.ss.util.CellUtil#translateUnicodeValues} instead.")]
        public static ICell TranslateUnicodeValues(ICell cell)
        {
            CellUtil.TranslateUnicodeValues(cell);
            return cell;
        }
        /// <summary>
        /// Translate color palette entries from the source to the destination sheet
        /// </summary>
        private static void RemapCellStyle(HSSFCellStyle stylish, Dictionary<short, short> paletteMap)
        {
            if (paletteMap.ContainsKey(stylish.BorderDiagonalColor))
            {
                stylish.BorderDiagonalColor = paletteMap[stylish.BorderDiagonalColor];
            }
            if (paletteMap.ContainsKey(stylish.BottomBorderColor))
            {
                stylish.BottomBorderColor = paletteMap[stylish.BottomBorderColor];
            }
            if (paletteMap.ContainsKey(stylish.FillBackgroundColor))
            {
                stylish.FillBackgroundColor = paletteMap[stylish.FillBackgroundColor];
            }
            if (paletteMap.ContainsKey(stylish.FillForegroundColor))
            {
                stylish.FillForegroundColor = paletteMap[stylish.FillForegroundColor];
            }
            if (paletteMap.ContainsKey(stylish.LeftBorderColor))
            {
                stylish.LeftBorderColor = paletteMap[stylish.LeftBorderColor];
            }
            if (paletteMap.ContainsKey(stylish.RightBorderColor))
            {
                stylish.RightBorderColor = paletteMap[stylish.RightBorderColor];
            }
            if (paletteMap.ContainsKey(stylish.TopBorderColor))
            {
                stylish.TopBorderColor = paletteMap[stylish.TopBorderColor];
            }
        }
        public static void CopyCell(HSSFCell oldCell, HSSFCell newCell, IDictionary<Int32, HSSFCellStyle> styleMap, Dictionary<short, short> paletteMap, Boolean keepFormulas)
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
                            HSSFCellStyle newCellStyle = (HSSFCellStyle)newCell.Sheet.Workbook.CreateCellStyle();
                            newCellStyle.CloneStyleFrom(oldCell.CellStyle);
                            RemapCellStyle(newCellStyle, paletteMap); //Clone copies as-is, we need to remap colors manually
                            newCell.CellStyle = newCellStyle;
                            //Clone of cell style always clones the font. This makes my life easier
                            IFont theFont = newCellStyle.GetFont(newCell.Sheet.Workbook);
                            if (theFont.Color > 0 && paletteMap.ContainsKey(theFont.Color))
                            {
                                theFont.Color = paletteMap[theFont.Color]; //Remap font color
                            }
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
                    HSSFRichTextString rts = oldCell.RichStringCellValue as HSSFRichTextString;
                    newCell.SetCellValue(rts);
                    if (rts != null)
                    {
                        for (int j = 0; j < rts.NumFormattingRuns; j++)
                        {
                            short fontIndex = rts.GetFontOfFormattingRun(j);
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
                            FontRecord fr = newCell.BoundWorkbook.CreateNewFont();
                            fr.CloneStyleFrom(oldCell.BoundWorkbook.GetFontRecordAt(fontIndex));
                            HSSFFont font = new HSSFFont((short)(newCell.BoundWorkbook.GetFontIndex(fr)), fr);
                            newCell.RichStringCellValue.ApplyFont(startIndex, endIndex, font);
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

    }
}