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

namespace NPOI.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.Util.ArrayExtensions;
    using NUnit.Framework;
    using static NPOI.SS.Formula.Functions.Frequency;

    /// <summary>
    /// Testcase for the function FREQUENCY(data, bins)
    /// </summary>
    /// <remarks>
    /// @author Yegor Kozlov
    /// </remarks>
    [TestFixture]
    public class TestFrequency
    {

        [Test]
        public void TestHistogram()
        {

            Assert.IsTrue(Arrays.Equals(new int[] { 3, 2, 2, 0, 1, 1 },
                    Histogram(
                            new double[] { 11, 12, 13, 21, 29, 36, 40, 58, 69 },
                            new double[] { 20, 30, 40, 50, 60 })
            ));

            Assert.IsTrue(Arrays.Equals(new int[] { 1, 1, 1, 1, 1, 0 },
                    Histogram(
                            new double[] { 20, 30, 40, 50, 60 },
                            new double[] { 20, 30, 40, 50, 60 })

            ));

            Assert.IsTrue(Arrays.Equals(new int[] { 2, 3 },
                    Histogram(
                            new double[] { 20, 30, 40, 50, 60 },
                            new double[] { 30 })

            ));
        }

        [Test]
        public void TestEvaluate()
        {
            IWorkbook wb = new HSSFWorkbook();
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            int[] data = {1, 1, 2, 3, 4, 4, 5, 7, 8, 9, 9, 11, 3, 5, 8};
            int[] bins = {3, 6, 9};
            ISheet sheet = wb.CreateSheet();
            IRow dataRow = sheet.CreateRow(0); // A1:O1
            for(int i = 0; i < data.Length; i++)
            {
                dataRow.CreateCell(i).SetCellValue(data[i]);
            }
            IRow binsRow = sheet.CreateRow(1);
            for(int i = 0; i < bins.Length; i++)
            { 
                // A2:C2
                binsRow.CreateCell(i).SetCellValue(bins[i]);
            }
            IRow fmlaRow = sheet.CreateRow(2);
            ICellRange<ICell> arrayFmla = sheet.SetArrayFormula("FREQUENCY(A1:O1,A2:C2)", CellRangeAddress.ValueOf("A3:A6"));
            ICell b3 = fmlaRow.CreateCell(1); // B3
            b3.SetCellFormula("COUNT(FREQUENCY(A1:O1,A2:C2))"); // frequency returns a vertical array of bins+1

            ICell c3 = fmlaRow.CreateCell(2);
            c3.SetCellFormula("SUM(FREQUENCY(A1:O1,A2:C2))"); // sum of the frequency bins should add up to the number of data values

            Assert.AreEqual(5, (int) evaluator.Evaluate(arrayFmla.FlattenedCells[0]).NumberValue);
            Assert.AreEqual(4, (int) evaluator.Evaluate(arrayFmla.FlattenedCells[1]).NumberValue);
            Assert.AreEqual(5, (int) evaluator.Evaluate(arrayFmla.FlattenedCells[2]).NumberValue);
            Assert.AreEqual(1, (int) evaluator.Evaluate(arrayFmla.FlattenedCells[3]).NumberValue);

            Assert.AreEqual(4, (int) evaluator.Evaluate(b3).NumberValue);
            Assert.AreEqual(15, (int) evaluator.Evaluate(c3).NumberValue);

        }
    }
}

