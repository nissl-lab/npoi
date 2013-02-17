/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System;
using NPOI.SS.Util;

namespace NPOI.SS.UserModel.Charts
{
    public class DataSources {

    private DataSources() {
    }

    public static IChartDataSource<T> FromArray<T>(T[] elements)
    {
        return new ArrayDataSource<T>(elements);
    }

    public static IChartDataSource<double> FromNumericCellRange(ISheet sheet, CellRangeAddress cellRangeAddress) {
        return new DoubleCellRangeDataSource(sheet, cellRangeAddress);
    }
    private class DoubleCellRangeDataSource: AbstractCellRangeDataSource<double>
    {
        public DoubleCellRangeDataSource(ISheet sheet, CellRangeAddress cellRangeAddress) 
            : base(sheet, cellRangeAddress)
        {
        }

        public override double GetPointAt(int index)
        {
            CellValue cellValue = GetCellValueAt(index);
            if (cellValue != null && cellValue.CellType == CellType.Numeric)
            {
                return cellValue.NumberValue;
            }
            else
            {
                return double.NaN;
            }
        }

        public override bool IsNumeric
        {
            get { return true; }
        }
    }
    public static IChartDataSource<String> FromStringCellRange(ISheet sheet, CellRangeAddress cellRangeAddress)
    {
        return new StringCellRangeDataSource(sheet, cellRangeAddress);
    }
    private class StringCellRangeDataSource : AbstractCellRangeDataSource<string> 
    {
        public StringCellRangeDataSource(ISheet sheet, CellRangeAddress cellRangeAddress) 
            : base(sheet, cellRangeAddress)
        {

        }

        public override string GetPointAt(int index)
        {
            CellValue cellValue = GetCellValueAt(index);
            if (cellValue != null && cellValue.CellType == CellType.String)
            {
                return cellValue.StringValue;
            }
            else
            {
                return null;
            }
        }

        public override bool IsNumeric
        {
            get { return false; }
        }
    }

    private class ArrayDataSource<T> : IChartDataSource<T> {

        private T[] elements;

        public ArrayDataSource(T[] elements) {
            this.elements = elements;
        }

        public int PointCount {
            get { return elements.Length; }
        }

        public T GetPointAt(int index) {
            return elements[index];
        }

        public bool IsReference {
            get { return false; }
        }

        public bool IsNumeric {
            get
            {
                //Class < ? > arrayComponentType = elements.getClass().getComponentType();
                //return (Number.class.isAssignableFrom(arrayComponentType))
                Type dt = typeof (double);
                foreach (object element in elements)
                {
                    if (!dt.IsInstanceOfType(element))
                        return false;
                }
                return true;
            }
        }

        public String FormulaString {
            get { throw new InvalidOperationException("Literal data source can not be expressed by reference."); }
        }
    }

    private abstract class AbstractCellRangeDataSource<T> : IChartDataSource<T> {
        private ISheet sheet;
        private CellRangeAddress cellRangeAddress;
        private int numOfCells;
        private IFormulaEvaluator evaluator;

        protected AbstractCellRangeDataSource(ISheet sheet, CellRangeAddress cellRangeAddress) {
            this.sheet = sheet;
            // Make copy since CellRangeAddress is mutable.
            this.cellRangeAddress = cellRangeAddress.Copy();
            this.numOfCells = this.cellRangeAddress.NumberOfCells;
            this.evaluator = sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
        }

        public int PointCount {
            get { return numOfCells; }
        }

        public bool IsReference {
            get { return true; }
        }

        public abstract T GetPointAt(int index);
        public abstract bool IsNumeric { get; }

        public String FormulaString {
            get { return cellRangeAddress.FormatAsString(sheet.SheetName, true); }
        }

        protected CellValue GetCellValueAt(int index) {
            if (index < 0 || index >= numOfCells) {
                throw new IndexOutOfRangeException("Index must be between 0 and " +
                        (numOfCells - 1) + " (inclusive), given: " + index);
            }
            int firstRow = cellRangeAddress.FirstRow;
            int firstCol = cellRangeAddress.FirstColumn;
            int lastCol = cellRangeAddress.LastColumn;
            int width = lastCol - firstCol + 1;
            int rowIndex = firstRow + index / width;
            int cellIndex = firstCol + index % width;
            IRow row = sheet.GetRow(rowIndex);
            return (row == null) ? null : evaluator.Evaluate(row.GetCell(cellIndex));
        }
    }
}
}