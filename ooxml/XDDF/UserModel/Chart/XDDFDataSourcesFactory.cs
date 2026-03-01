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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NPOI.OpenXmlFormats.Dml.Chart;

    /// <summary>
    /// Class <c>XDDFDataSourcesFactory</c> is a factory for <see cref="XDDFDataSource"/> instances.
    /// </summary>
    public class XDDFDataSourcesFactory
    {

        private XDDFDataSourcesFactory()
        {
        }

        public static XDDFCategoryDataSource FromDataSource(CT_AxDataSource categoryDS)
        {
            if(categoryDS.strRef == null)
            {
                return new XDDFCategoryDataSourceNumeric(categoryDS);
            }
            else
            {
                return new XDDFCategoryDataSourceString(categoryDS);
            }
        }

        public class XDDFCategoryDataSourceNumeric : XDDFCategoryDataSource
        {
            private CT_NumData category;
            private CT_AxDataSource categoryDS;
            public XDDFCategoryDataSourceNumeric(CT_AxDataSource categoryDS)
            {
                this.categoryDS = categoryDS;
                category = (CT_NumData) categoryDS.numRef.numCache.Copy();
            }
            public bool IsNumeric => true;
            public string Formula => categoryDS.numRef.f;
            public int PointCount => (int) category.ptCount.val;

            public bool IsReference => true;

            public int ColIndex => 0;

            public string DataRangeReference => Formula;

            public string GetPointAt(int index)
            {
                return category.GetPtArray(index).v;
            }
        }

        public class XDDFCategoryDataSourceString : XDDFCategoryDataSource
        {
            private CT_StrData category;
            private CT_AxDataSource categoryDS;
            public XDDFCategoryDataSourceString(CT_AxDataSource categoryDS)
            {
                category = (CT_StrData) categoryDS.strRef.strCache.Copy();
                this.categoryDS = categoryDS;
            }
            public string Formula => categoryDS.strRef.f;

            public int PointCount => (int) category.ptCount.val;

            public bool IsReference => true;

            public bool IsNumeric => false;

            public int ColIndex => 0;

            public string DataRangeReference => Formula;

            public string GetPointAt(int index)
            {
                return category.GetPtArray(index).v;
            }
        }

        public static IXDDFNumericalDataSource<Double> FromDataSource(CT_NumDataSource valuesDS)
        {
            return new XDDFNumericalDataSource(valuesDS);
        }

        public class XDDFNumericalDataSource : IXDDFNumericalDataSource<Double>
        {
            private CT_NumDataSource valuesDS;
            private CT_NumData values;
            private string formatCode;
            public XDDFNumericalDataSource(CT_NumDataSource valuesDS)
            {
                this.valuesDS = valuesDS;
                values = (CT_NumData) valuesDS.numRef.numCache.Copy();
                formatCode = values.IsSetFormatCode() ? values.formatCode : null;
            }
            public string Formula => valuesDS.numRef.f;
            public string FormatCode
            {
                get => formatCode;
                set => this.formatCode = value;
            }
            public bool IsNumeric => true;
            public bool IsReference => true;
            public int PointCount => (int) values.ptCount.val;
            public Double GetPointAt(int index)
            {
                return double.Parse(values.GetPtArray(index).v);
            }
            public string DataRangeReference => valuesDS.numRef.f;
            public int ColIndex => 0;
        }

        public static IXDDFNumericalDataSource<T> FromArray<T>(T[] elements, string dataRange)
        {
            return new NumericalArrayDataSource<T>(elements, dataRange);
        }

        public static XDDFCategoryDataSource FromArray(String[] elements, string dataRange)
        {
            return new StringArrayDataSource(elements, dataRange);
        }

        public static IXDDFNumericalDataSource<T> FromArray<T>(T[] elements, string dataRange, int col)
        {
            return new NumericalArrayDataSource<T>(elements, dataRange, col);
        }

        public static XDDFCategoryDataSource FromArray(String[] elements, string dataRange, int col)
        {
            return new StringArrayDataSource(elements, dataRange, col);
        }

        public static IXDDFNumericalDataSource<Double> FromNumericCellRange(XSSFSheet sheet,
                CellRangeAddress cellRangeAddress)
        {
            return new NumericalCellRangeDataSource(sheet, cellRangeAddress);
        }

        public static XDDFCategoryDataSource FromStringCellRange(XSSFSheet sheet, CellRangeAddress cellRangeAddress)
        {
            return new StringCellRangeDataSource(sheet, cellRangeAddress);
        }

        private abstract class AbstractArrayDataSource<T> : IXDDFDataSource<T>
        {
            private  T[] elements;
            private string dataRange;
            private int col = 0;

            public AbstractArrayDataSource(T[] elements, string dataRange)
            {
                this.elements = (T[]) elements.Clone();
                this.dataRange = dataRange;
            }

            public AbstractArrayDataSource(T[] elements, string dataRange, int col)
            {
                this.elements = (T[]) elements.Clone();
                this.dataRange = dataRange;
                this.col = col;
            }
            public int PointCount
            {
                get
                {
                    return elements.Length;
                }

            }
            public T GetPointAt(int index)
            {
                return elements[index];
            }
            public bool IsReference
            {
                get { return dataRange != null; }
            }
            public bool IsNumeric
            {
                get
                {
                    return Number.IsNumber(elements.GetType().GetElementType());
                }

            }
            public string DataRangeReference
            {
                get
                {
                    if(dataRange == null)
                    {
                        throw new InvalidOperationException("Literal data source can not be expressed by reference.");
                    }
                    else
                    {
                        return dataRange;
                    }
                }
            }
            public int ColIndex
            {
                get { return col; }
            }

            public string Formula { get => string.Empty; }
        }

        private class NumericalArrayDataSource<T> :
                AbstractArrayDataSource<T>, IXDDFNumericalDataSource<T>
        {
            private string formatCode;

            public NumericalArrayDataSource(T[] elements, string dataRange)
                    : base(elements, dataRange)
            {

            }

            public NumericalArrayDataSource(T[] elements, string dataRange, int col)
                    : base(elements, dataRange, col)
            {

            }
            public string Formula
            {
                get
                {
                    return DataRangeReference;
                }
            }
            public string FormatCode
            {
                get
                {
                    return formatCode;
                }
                set
                {
                    this.formatCode = value;
                }
            }

        }

        private class StringArrayDataSource : AbstractArrayDataSource<string>
                , XDDFCategoryDataSource
        {
            public StringArrayDataSource(String[] elements, string dataRange)
                    : base(elements, dataRange)
            {

            }

            public StringArrayDataSource(String[] elements, string dataRange, int col)
                    : base(elements, dataRange, col)
            {

            }
            public string Formula
            {
                get
                {
                    return DataRangeReference;
                }

            }
        }

        private abstract class AbstractCellRangeDataSource<T> : IXDDFDataSource<T>
        {
            private  XSSFSheet sheet;
            private  CellRangeAddress cellRangeAddress;
            private  int numOfCells;
            private XSSFFormulaEvaluator evaluator;

            internal AbstractCellRangeDataSource(XSSFSheet sheet, CellRangeAddress cellRangeAddress)
            {
                this.sheet = sheet;
                // Make copy since CellRangeAddress is mutable.
                this.cellRangeAddress = cellRangeAddress.Copy();
                this.numOfCells = this.cellRangeAddress.NumberOfCells;
                this.evaluator = sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator() as XSSFFormulaEvaluator;
            }
            public int PointCount
            {
                get { return numOfCells; }
            }
            public bool IsReference
            {
                get { return true; }
            }
            public int ColIndex
            {
                get { return cellRangeAddress.FirstColumn; }
            }
            public string DataRangeReference
            {
                get { return cellRangeAddress.FormatAsString(sheet.SheetName, true); }
            }

            public bool IsNumeric => throw new NotImplementedException();

            public string Formula => throw new NotImplementedException();

            public T GetPointAt(int index)
            {
                throw new NotImplementedException();
            }

            protected CellValue GetCellValueAt(int index)
            {
                if(index < 0 || index >= numOfCells)
                {
                    throw new IndexOutOfRangeException(
                            "Index must be between 0 and " + (numOfCells - 1) + " (inclusive), given: " + index);
                }
                int firstRow = cellRangeAddress.FirstRow;
                int firstCol = cellRangeAddress.FirstColumn;
                int lastCol = cellRangeAddress.LastColumn;
                int width = lastCol - firstCol + 1;
                int rowIndex = firstRow + index / width;
                int cellIndex = firstCol + index % width;
                XSSFRow row = sheet.GetRow(rowIndex) as XSSFRow;
                return (row == null) ? null : evaluator.Evaluate(row.GetCell(cellIndex));
            }
        }

        private class NumericalCellRangeDataSource : AbstractCellRangeDataSource<Double>
                , IXDDFNumericalDataSource<Double>
        {
            internal NumericalCellRangeDataSource(XSSFSheet sheet, CellRangeAddress cellRangeAddress)
                    : base(sheet, cellRangeAddress)
            {

            }
            public string Formula
            {
                get
                {
                    return DataRangeReference;
                }
            }

            private string formatCode;
            public string FormatCode
            {
                get { return formatCode; }
                set { this.formatCode = value; }
            }

            public Double GetPointAt(int index)
            {
                CellValue cellValue = GetCellValueAt(index);
                if(cellValue != null && cellValue.CellType == CellType.Numeric)
                {
                    return cellValue.NumberValue;
                }
                else
                {
                    return Double.NaN;
                }
            }
            public bool IsNumeric
            {
                get { return true; }
            }
        }

        private class StringCellRangeDataSource : AbstractCellRangeDataSource<String>
                , XDDFCategoryDataSource
        {
            internal StringCellRangeDataSource(XSSFSheet sheet, CellRangeAddress cellRangeAddress)
                    : base(sheet, cellRangeAddress)
            {

            }
            public string Formula
            {
                get { return DataRangeReference; }
            }
            public string GetPointAt(int index)
            {
                CellValue cellValue = GetCellValueAt(index);
                if(cellValue != null && cellValue.CellType == CellType.String)
                {
                    return cellValue.StringValue;
                }
                else
                {
                    return null;
                }
            }
            public bool IsNumeric
            {
                get { return false; }
            }
        }
    }
}


