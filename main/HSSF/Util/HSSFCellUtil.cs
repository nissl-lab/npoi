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
        /// <summary>
        /// Translate color palette entries from the source to the destination sheet
        /// </summary>
        private static void RemapCellStyle(HSSFCellStyle stylish, Dictionary<short, short> paletteMap)
        {
            if (paletteMap.TryGetValue(stylish.BorderDiagonalColor, out short value))
            {
                stylish.BorderDiagonalColor = value;
            }
            if (paletteMap.TryGetValue(stylish.BottomBorderColor, out short value1))
            {
                stylish.BottomBorderColor = value1;
            }
            if (paletteMap.TryGetValue(stylish.FillBackgroundColor, out short value2))
            {
                stylish.FillBackgroundColor = value2;
            }
            if (paletteMap.TryGetValue(stylish.FillForegroundColor, out short value3))
            {
                stylish.FillForegroundColor = value3;
            }
            if (paletteMap.TryGetValue(stylish.LeftBorderColor, out short value4))
            {
                stylish.LeftBorderColor = value4;
            }
            if (paletteMap.TryGetValue(stylish.RightBorderColor, out short value5))
            {
                stylish.RightBorderColor = value5;
            }
            if (paletteMap.TryGetValue(stylish.TopBorderColor, out short value6))
            {
                stylish.TopBorderColor = value6;
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
                        if (styleMap.TryGetValue(styleHashCode, out HSSFCellStyle value))
                        {
                            newCell.CellStyle = value;
                        }
                        else
                        {
                            HSSFCellStyle newCellStyle = (HSSFCellStyle)newCell.Sheet.Workbook.CreateCellStyle();
                            newCellStyle.CloneStyleFrom(oldCell.CellStyle);
                            RemapCellStyle(newCellStyle, paletteMap); //Clone copies as-is, we need to remap colors manually
                            newCell.CellStyle = newCellStyle;
                            //Clone of cell style always clones the font. This makes my life easier
                            IFont theFont = newCellStyle.GetFont(newCell.Sheet.Workbook);
                            if (theFont.Color > 0 && paletteMap.TryGetValue(theFont.Color, out short value1))
                            {
                                theFont.Color = value1; //Remap font color
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