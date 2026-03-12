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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XDDF.UserModel.Chart
{
    using NPOI.SS.Util;
    using NPOI.XDDF.UserModel.Chart;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Tests for <see cref="XDDFDataSourcesFactory"/>.
    /// </summary>
    [TestFixture]
    public class TestXDDFDataSourcesFactory
    {

        private static  object[][] numericCells = [
            [0.0,      1.0,       2.0,     3.0,      4.0],
            [0.0, "=B1*2",  "=C1*2", "=D1*2", "=E1*2"]
        ];

        private static  object[][] stringCells = [
            [  1,    2,    3,   4,    5],
            ["A", "B", "C", "D", "E"]
        ];

        private static  object[][] mixedCells = [
            [1.0, "2.0", 3.0, "4.0", 5.0, "6.0"]
        ];
        [Test]
        public void TestNumericArrayDataSource()
        {
            Double[] doubles = new Double[]{1.0, 2.0, 3.0, 4.0, 5.0};
            IXDDFDataSource<Double> doubleDataSource = XDDFDataSourcesFactory.FromArray(doubles, null);
            ClassicAssert.IsTrue(doubleDataSource.IsNumeric);
            ClassicAssert.IsFalse(doubleDataSource.IsReference);
            assertDataSourceIsEqualToArray(doubleDataSource, doubles);
        }
        [Test]
        public void TestStringArrayDataSource()
        {
            String[] strings = new String[]{"one", "two", "three", "four", "five"};
            IXDDFDataSource<String> stringDataSource = XDDFDataSourcesFactory.FromArray(strings, null);
            ClassicAssert.IsFalse(stringDataSource.IsNumeric);
            ClassicAssert.IsFalse(stringDataSource.IsReference);
            assertDataSourceIsEqualToArray(stringDataSource, strings);
        }
        [Test]
        public void TestNumericCellDataSource()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet) new SheetBuilder(wb, numericCells).Build();
            CellRangeAddress numCellRange = CellRangeAddress.ValueOf("A2:E2");
            IXDDFDataSource<Double> numDataSource = XDDFDataSourcesFactory.FromNumericCellRange(sheet, numCellRange);
            ClassicAssert.IsTrue(numDataSource.IsReference);
            ClassicAssert.IsTrue(numDataSource.IsNumeric);
            ClassicAssert.AreEqual(numericCells[0].Length, numDataSource.PointCount);
            for(int i = 0; i < numericCells[0].Length; ++i)
            {
                ClassicAssert.AreEqual(((Double) numericCells[0][i]) * 2,
                        numDataSource.GetPointAt(i), 0.00001);
            }
        }
        [Test]
        public void TestStringCellDataSource()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet) new SheetBuilder(wb, stringCells).Build();
            CellRangeAddress numCellRange = CellRangeAddress.ValueOf("A2:E2");
            IXDDFDataSource<String> numDataSource = XDDFDataSourcesFactory.FromStringCellRange(sheet, numCellRange);
            ClassicAssert.IsTrue(numDataSource.IsReference);
            ClassicAssert.IsFalse(numDataSource.IsNumeric);
            ClassicAssert.AreEqual(numericCells[0].Length, numDataSource.PointCount);
            for(int i = 0; i < stringCells[1].Length; ++i)
            {
                ClassicAssert.AreEqual(stringCells[1][i], numDataSource.GetPointAt(i));
            }
        }
        [Test]
        public void TestMixedCellDataSource()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet) new SheetBuilder(wb, mixedCells).Build();
            CellRangeAddress mixedCellRange = CellRangeAddress.ValueOf("A1:F1");
            IXDDFDataSource<String> strDataSource = XDDFDataSourcesFactory.FromStringCellRange(sheet, mixedCellRange);
            IXDDFDataSource<Double> numDataSource = XDDFDataSourcesFactory.FromNumericCellRange(sheet, mixedCellRange);
            for(int i = 0; i < mixedCells[0].Length; ++i)
            {
                if(i % 2 == 0)
                {
                    ClassicAssert.IsNull(strDataSource.GetPointAt(i));
                    ClassicAssert.AreEqual(((Double) mixedCells[0][i]),
                            numDataSource.GetPointAt(i), 0.00001);
                }
                else
                {
                    ClassicAssert.IsNaN(numDataSource.GetPointAt(i));
                    ClassicAssert.AreEqual(mixedCells[0][i], strDataSource.GetPointAt(i));
                }
            }
        }
        [Test]
        public void TestIOBExceptionOnInvalidIndex()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet) new SheetBuilder(wb, numericCells).Build();
            CellRangeAddress rangeAddress = CellRangeAddress.ValueOf("A2:E2");
            IXDDFDataSource<Double> numDataSource = XDDFDataSourcesFactory.FromNumericCellRange(sheet, rangeAddress);
            IndexOutOfRangeException exception = null;
            try
            {
                numDataSource.GetPointAt(-1);
            }
            catch(IndexOutOfRangeException e)
            {
                exception = e;
            }
            ClassicAssert.IsNotNull(exception);

            exception = null;
            try
            {
                numDataSource.GetPointAt(numDataSource.PointCount);
            }
            catch(IndexOutOfRangeException e)
            {
                exception = e;
            }
            ClassicAssert.IsNotNull(exception);
        }

        private void assertDataSourceIsEqualToArray<T>(IXDDFDataSource<T> ds, T[] array)
        {
            ClassicAssert.AreEqual(ds.PointCount, array.Length);
            for(int i = 0; i < array.Length; ++i)
            {
                ClassicAssert.AreEqual(ds.GetPointAt(i), array[i]);
            }
        }
    }
}
