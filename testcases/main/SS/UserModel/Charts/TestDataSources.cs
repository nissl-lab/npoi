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
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NUnit.Framework;

namespace TestCases.SS.UserModel.Charts
{
    /**
    * Tests for {@link org.apache.poi.ss.usermodel.charts.DataSources}.
    *
    * @author Roman Kashitsyn
    */
    [TestFixture]
    public class TestDataSources
    {

        private static readonly Object[][] numericCells =
            {
                new object[] {0.0, 1.0, 2.0, 3.0, 4.0},
                new object[] {0.0, "=B1*2", "=C1*2", "=D1*2", "=E1*2"}
            };

        private static readonly Object[][] stringCells =
            {
                new object[]{1, 2, 3, 4, 5},
                new object[]{"A", "B", "C", "D", "E"}
            };

        private static readonly Object[][] mixedCells =
            {
                new object[]{1.0, "2.0", 3.0, "4.0", 5.0, "6.0"}
            };
        [Test]
        public void TestNumericArrayDataSource()
        {
            Double[] doubles = new Double[] {1.0, 2.0, 3.0, 4.0, 5.0};
            IChartDataSource<Double> doubleDataSource = DataSources.FromArray(doubles);
            Assert.IsTrue(doubleDataSource.IsNumeric);
            Assert.IsFalse(doubleDataSource.IsReference);
            AssertDataSourceIsEqualToArray(doubleDataSource, doubles);
        }
        [Test]
        public void TestStringArrayDataSource()
        {
            String[] strings = new String[] {"one", "two", "three", "four", "five"};
            IChartDataSource<String> stringDataSource = DataSources.FromArray(strings);
            Assert.IsFalse(stringDataSource.IsNumeric);
            Assert.IsFalse(stringDataSource.IsReference);
            AssertDataSourceIsEqualToArray(stringDataSource, strings);
        }
        [Test]
        public void TestNumericCellDataSource()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, numericCells).Build();
            CellRangeAddress numCellRange = CellRangeAddress.ValueOf("A2:E2");
            IChartDataSource<double> numDataSource = DataSources.FromNumericCellRange(sheet, numCellRange);
            Assert.IsTrue(numDataSource.IsReference);
            Assert.IsTrue(numDataSource.IsNumeric);
            Assert.AreEqual(numericCells[0].Length, numDataSource.PointCount);
            for (int i = 0; i < numericCells[0].Length; ++i)
            {
                Assert.AreEqual(((double) numericCells[0][i])*2,
                                numDataSource.GetPointAt(i), 0.00001);
            }
        }
        [Test]
        public void TestStringCellDataSource()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, stringCells).Build();
            CellRangeAddress numCellRange = CellRangeAddress.ValueOf("A2:E2");
            IChartDataSource<String> numDataSource = DataSources.FromStringCellRange(sheet, numCellRange);
            Assert.IsTrue(numDataSource.IsReference);
            Assert.IsFalse(numDataSource.IsNumeric);
            Assert.AreEqual(numericCells[0].Length, numDataSource.PointCount);
            for (int i = 0; i < stringCells[1].Length; ++i)
            {
                Assert.AreEqual(stringCells[1][i], numDataSource.GetPointAt(i));
            }
        }
        [Test]
        public void TestMixedCellDataSource()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, mixedCells).Build();
            CellRangeAddress mixedCellRange = CellRangeAddress.ValueOf("A1:F1");
            IChartDataSource<String> strDataSource = DataSources.FromStringCellRange(sheet, mixedCellRange);
            IChartDataSource<double> numDataSource = DataSources.FromNumericCellRange(sheet, mixedCellRange);
            for (int i = 0; i < mixedCells[0].Length; ++i)
            {
                if (i%2 == 0)
                {
                    Assert.IsNull(strDataSource.GetPointAt(i));
                    Assert.AreEqual(((double) mixedCells[0][i]),
                                    numDataSource.GetPointAt(i), 0.00001);
                }
                else
                {
                    Assert.IsNaN(numDataSource.GetPointAt(i));
                    Assert.AreEqual(mixedCells[0][i], strDataSource.GetPointAt(i));
                }
            }
        }
        [Test]
        public void TestIobExceptionOnInvalidIndex()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, numericCells).Build();
            CellRangeAddress rangeAddress = CellRangeAddress.ValueOf("A2:E2");
            IChartDataSource<double> numDataSource = DataSources.FromNumericCellRange(sheet, rangeAddress);
            IndexOutOfRangeException exception = null;
            try
            {
                numDataSource.GetPointAt(-1);
            }
            catch (IndexOutOfRangeException e)
            {
                exception = e;
            }
            Assert.IsNotNull(exception);

            exception = null;
            try
            {
                numDataSource.GetPointAt(numDataSource.PointCount);
            }
            catch (IndexOutOfRangeException e)
            {
                exception = e;
            }
            Assert.IsNotNull(exception);
        }

        private void AssertDataSourceIsEqualToArray<T>(IChartDataSource<T> ds, T[] array)
        {
            Assert.AreEqual(ds.PointCount, array.Length);
            for (int i = 0; i < array.Length; ++i)
            {
                Assert.AreEqual(ds.GetPointAt(i), array[i]);
            }
        }
    }

}