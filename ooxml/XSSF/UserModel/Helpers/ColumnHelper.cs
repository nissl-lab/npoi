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

using System.Collections.Generic;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.Util;
namespace NPOI.XSSF.UserModel.Helpers
{
    /**
     * Helper class for dealing with the Column Settings on
     *  a CT_Worksheet (the data part of a sheet).
     * Note - within POI, we use 0 based column indexes, but
     *  the column defInitions in the XML are 1 based!
     */
    public class ColumnHelper
    {

        private CT_Worksheet worksheet;
        private CT_Cols newCols;

        public ColumnHelper(CT_Worksheet worksheet)
        {

            this.worksheet = worksheet;
            CleanColumns();
        }

       public void CleanColumns()
        {
            this.newCols = new CT_Cols();
            List<CT_Cols> colsArray = worksheet.GetColsArray();
            if (null != colsArray)
            {
                int i = 0;
                for (i = 0; i < colsArray.Count; i++)
                {
                    CT_Cols cols = colsArray[i];
                    List<CT_Col> colArray = cols.GetColArray();
                    for (int y = 0; y < colArray.Count; y++)
                    {
                        CT_Col col = colArray[y];
                        newCols = AddCleanColIntoCols(newCols, col);
                    }
                }
                for (int y = i - 1; y >= 0; y--)
                {
                    worksheet.RemoveCols(y);
                }
            }
            worksheet.AddNewCols();
            worksheet.SetColsArray(0, newCols);
        }

        //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
        public static void SortColumns(CT_Cols newCols)
        {
            List<CT_Col> colArray = newCols.GetColArray();
            colArray.Sort(new CTColComparator());
            newCols.SetColArray(colArray);
        }

        public CT_Col CloneCol(CT_Cols cols, CT_Col col)
        {
            CT_Col newCol = cols.AddNewCol();
            newCol.min = (uint)(col.min);
            newCol.max = (uint)(col.max);
            SetColumnAttributes(col, newCol);
            return newCol;
        }

        /**
         * Returns the Column at the given 0 based index
         */
        public CT_Col GetColumn(long index, bool splitColumns)
        {
            return GetColumn1Based(index + 1, splitColumns);
        }
        /**
         * Returns the Column at the given 1 based index.
         * POI default is 0 based, but the file stores
         *  as 1 based.
         */
        public CT_Col GetColumn1Based(long index1, bool splitColumns)
        {
            CT_Cols colsArray = worksheet.GetColsArray(0);
            for (int i = 0; i < colsArray.sizeOfColArray(); i++)
            {
                CT_Col colArray = colsArray.GetColArray(i);
                if (colArray.min <= index1 && colArray.max >= index1)
                {
                    if (splitColumns)
                    {
                        if (colArray.min < index1)
                        {
                            insertCol(colsArray, colArray.min, (index1 - 1), new CT_Col[] { colArray });
                        }
                        if (colArray.max > index1)
                        {
                            insertCol(colsArray, (index1 + 1), colArray.max, new CT_Col[] { colArray });
                        }
                        colArray.min = (uint)(index1);
                        colArray.max = (uint)(index1);
                    }
                    return colArray;
                }
            }
            return null;
        }

        public CT_Cols AddCleanColIntoCols(CT_Cols cols, CT_Col col)
        {
            bool colOverlaps = false;
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                CT_Col ithCol = cols.GetColArray(i);
                long[] range1 = { ithCol.min, ithCol.max };
                long[] range2 = { col.min, col.max };
                long[] overlappingRange = NumericRanges.GetOverlappingRange(range1,
                        range2);
                int overlappingType = NumericRanges.GetOverlappingType(range1,
                        range2);
                // different behavior required for each of the 4 different
                // overlapping types
                if (overlappingType == NumericRanges.OVERLAPS_1_MINOR)
                {
                    ithCol.max = (uint)(overlappingRange[0] - 1);
                    CT_Col rangeCol = insertCol(cols, overlappingRange[0],
                            overlappingRange[1], new CT_Col[] { ithCol, col });
                    i++;
                    CT_Col newCol = insertCol(cols, (overlappingRange[1] + 1), col
                            .max, new CT_Col[] { col });
                    i++;
                }
                else if (overlappingType == NumericRanges.OVERLAPS_2_MINOR)
                {
                    ithCol.min = (uint)(overlappingRange[1] + 1);
                    CT_Col rangeCol = insertCol(cols, overlappingRange[0],
                            overlappingRange[1], new CT_Col[] { ithCol, col });
                    i++;
                    CT_Col newCol = insertCol(cols, col.min,
                            (overlappingRange[0] - 1), new CT_Col[] { col });
                    i++;
                }
                else if (overlappingType == NumericRanges.OVERLAPS_2_WRAPS)
                {
                    SetColumnAttributes(col, ithCol);
                    if (col.min != ithCol.min)
                    {
                        CT_Col newColBefore = insertCol(cols, col.min, (ithCol
                                .min - 1), new CT_Col[] { col });
                        i++;
                    }
                    if (col.max != ithCol.max)
                    {
                        CT_Col newColAfter = insertCol(cols, (ithCol.max + 1),
                                col.max, new CT_Col[] { col });
                        i++;
                    }
                }
                else if (overlappingType == NumericRanges.OVERLAPS_1_WRAPS)
                {
                    if (col.min != ithCol.min)
                    {
                        CT_Col newColBefore = insertCol(cols, ithCol.min, (col
                                .min - 1), new CT_Col[] { ithCol });
                        i++;
                    }
                    if (col.max != ithCol.max)
                    {
                        CT_Col newColAfter = insertCol(cols, (col.max + 1),
                                ithCol.max, new CT_Col[] { ithCol });
                        i++;
                    }
                    ithCol.min = (uint)(overlappingRange[0]);
                    ithCol.max = (uint)(overlappingRange[1]);
                    SetColumnAttributes(col, ithCol);
                }
                if (overlappingType != NumericRanges.NO_OVERLAPS)
                {
                    colOverlaps = true;
                }
            }
            if (!colOverlaps)
            {
                CT_Col newCol = CloneCol(cols, col);
            }
            SortColumns(cols);
            return cols;
        }

        /*
         * Insert a new CT_Col at position 0 into cols, Setting min=min, max=max and
         * copying all the colsWithAttributes array cols attributes into newCol
         */
        private CT_Col insertCol(CT_Cols cols, long min, long max,
            CT_Col[] colsWithAttributes)
        {
            if (!columnExists(cols, min, max))
            {
                CT_Col newCol = cols.InsertNewCol(0);
                newCol.min = (uint)(min);
                newCol.max = (uint)(max);
                foreach (CT_Col col in colsWithAttributes)
                {
                    SetColumnAttributes(col, newCol);
                }
                return newCol;
            }
            return null;
        }

        /**
         * Does the column at the given 0 based index exist
         *  in the supplied list of column defInitions?
         */
        public bool columnExists(CT_Cols cols, long index)
        {
            return columnExists1Based(cols, index + 1);
        }
        private bool columnExists1Based(CT_Cols cols, long index1)
        {
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                if (cols.GetColArray(i).min == index1)
                {
                    return true;
                }
            }
            return false;
        }

        public void SetColumnAttributes(CT_Col fromCol, CT_Col toCol)
        {

            if (fromCol.IsSetBestFit())
            {
                toCol.bestFit = (fromCol.bestFit);
            }
            if (fromCol.IsSetCustomWidth())
            {
                toCol.customWidth = (fromCol.customWidth);
            }
            if (fromCol.IsSetHidden()) 
            {
                toCol.hidden = (fromCol.hidden);
            }
            if (fromCol.IsSetStyle())
            {
                toCol.style = (fromCol.style);
                toCol.styleSpecified = true;
            }
            if (fromCol.IsSetWidth())
            {
                toCol.width = (fromCol.width);
            }
            if (fromCol.IsSetCollapsed())
            {
                toCol.collapsed = (fromCol.collapsed);
            }
            if (fromCol.IsSetPhonetic())
            {
                toCol.phonetic = (fromCol.phonetic);
            }
            if (fromCol.IsSetOutlineLevel())
            {
                toCol.outlineLevel = (fromCol.outlineLevel);
            }
            if (fromCol.IsSetCollapsed())
            {
                toCol.collapsed = fromCol.collapsed;
            }
        }

        public void SetColBestFit(long index, bool bestFit)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, false);
            col.bestFit = (bestFit);
        }
        public void SetCustomWidth(long index, bool width)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.customWidth = (width);
        }

        public void SetColWidth(long index, double width)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.width = (width);
        }

        public void SetColHidden(long index, bool hidden)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.hidden = (hidden);
        }

        /**
         * Return the CT_Col at the given (0 based) column index,
         *  creating it if required.
         */
        internal CT_Col GetOrCreateColumn1Based(long index1, bool splitColumns)
        {
            CT_Col col = GetColumn1Based(index1, splitColumns);
            if (col == null)
            {
                col = worksheet.GetColsArray(0).AddNewCol();
                col.min = (uint)(index1);
                col.max = (uint)(index1);
            }
            return col;
        }

        public void SetColDefaultStyle(long index, ICellStyle style)
        {
            SetColDefaultStyle(index, style.Index);
        }

        public void SetColDefaultStyle(long index, int styleId)
        {
            CT_Col col = GetOrCreateColumn1Based(index + 1, true);
            col.style = (uint)styleId;
            col.styleSpecified = true;
        }

        // Returns -1 if no column is found for the given index
        public int GetColDefaultStyle(long index)
        {
            if (GetColumn(index, false) != null)
            {
                return (int)GetColumn(index, false).style;
            }
            return -1;
        }

        private bool columnExists(CT_Cols cols, long min, long max)
        {
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                if (cols.GetColArray(i).min == min && cols.GetColArray(i).max == max)
                {
                    return true;
                }
            }
            return false;
        }

        public int GetIndexOfColumn(CT_Cols cols, CT_Col col)
        {
            for (int i = 0; i < cols.sizeOfColArray(); i++)
            {
                if (cols.GetColArray(i).min == col.min && cols.GetColArray(i).max == col.max)
                {
                    return i;
                }
            }
            return -1;
        }
    }

}
