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
using System;
using NPOI.SS.UserModel;
using NPOI.Util;

namespace NPOI.SS.Util
{

    /**
     * Class {@code SheetBuilder} provides an easy way of building workbook sheets
     * from 2D array of Objects. It can be used in test cases to improve code
     * readability or in Swing applications with tables.
     *
     * @author Roman Kashitsyn
     */
    public class SheetBuilder
    {

        private IWorkbook workbook;
        private Object[][] cells;
        private bool shouldCreateEmptyCells = false;
        private String sheetName = null;
        public SheetBuilder(IWorkbook workbook, Object[][] cells)
        {
            this.workbook = workbook;
            this.cells = cells;
        }

        /**
         * Returns {@code true} if null array elements should be treated as empty
         * cells.
         *
         * @return {@code true} if null objects should be treated as empty cells
         *         and {@code false} otherwise
         */
        public bool GetCreateEmptyCells()
        {
            return shouldCreateEmptyCells;
        }

        /**
         * Specifies if null array elements should be treated as empty cells.
         *
         * @param shouldCreateEmptyCells {@code true} if null array elements should be
         *                               treated as empty cells
         * @return {@code this}
         */
        public SheetBuilder SetCreateEmptyCells(bool shouldCreateEmptyCells)
        {
            this.shouldCreateEmptyCells = shouldCreateEmptyCells;
            return this;
        }
        /**
         * Specifies name of the sheet to build. If not specified, default name (provided by
         * workbook) will be used instead.
         * @param sheetName sheet name to use
         * @return {@code this}
         */
        public SheetBuilder SetSheetName(String sheetName)
        {
            this.sheetName = sheetName;
            return this;
        }
        /**
          * Builds sheet from parent workbook and 2D array with cell
          * values. Creates rows anyway (even if row contains only null
          * cells), creates cells if either corresponding array value is not
          * null or createEmptyCells property is true.
          * The conversion is performed in the following way:
          * <p/>
          * <ul>
          * <li>Numbers become numeric cells.</li>
          * <li><code>java.util.Date</code> or <code>java.util.Calendar</code>
          * instances become date cells.</li>
          * <li>String with leading '=' char become formulas (leading '='
          * will be truncated).</li>
          * <li>Other objects become strings via <code>Object.toString()</code>
          * method call.</li>
          * </ul>
          *
          * @return newly created sheet
          */
        public ISheet Build()
        {
            ISheet sheet = (sheetName == null) ? workbook.CreateSheet() : workbook.CreateSheet(sheetName);
            IRow currentRow = null;
            ICell currentCell = null;

            for (int rowIndex = 0; rowIndex < cells.Length; ++rowIndex)
            {
                Object[] rowArray = cells[rowIndex];
                currentRow = sheet.CreateRow(rowIndex);

                for (int cellIndex = 0; cellIndex < rowArray.Length; ++cellIndex)
                {
                    Object cellValue = rowArray[cellIndex];
                    if (cellValue != null || shouldCreateEmptyCells)
                    {
                        currentCell = currentRow.CreateCell(cellIndex);
                        SetCellValue(currentCell, cellValue);
                    }
                }
            }
            return sheet;
        }

        /**
         * Sets the cell value using object type information.
         * @param cell cell to change
         * @param value value to set
         */
        private void SetCellValue(ICell cell, Object value)
        {
            if (value == null || cell == null)
            {
                return;
            }
            //else if (value is Number)
            else if (Number.IsNumber(value))
            {
                double val;
                double.TryParse(value.ToString(), out val);
                cell.SetCellValue(val);
            }
            else if (value is DateTime)
            {
                cell.SetCellValue((DateTime)value);
                //} else if (value is Calendar) {
                //    cell.SetCellValue((Calendar) value);
            }
            else if (IsFormulaDefinition(value))
            {
                cell.CellFormula = (GetFormula(value));
            }
            else
            {
                cell.SetCellValue(value.ToString());
            }
        }

        private bool IsFormulaDefinition(Object obj)
        {
            if (obj is String)
            {
                String str = (String)obj;
                if (str.Length < 2)
                {
                    return false;
                }
                else
                {
                    return ((String)obj)[0] == '=';
                }
            }
            else
            {
                return false;
            }
        }

        private String GetFormula(Object obj)
        {
            return ((String)obj).Substring(1);
        }
    }
}