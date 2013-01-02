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

using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using NUnit.Framework;
using TestCases.HSSF;

namespace TestCases.SS.Formula.Functions
{
    /**
     * Test cases for RANK()
     */
    [TestFixture]
    public class TestRank
    {
        [Test]
        public void TestFromFile()
        {

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("rank.xls");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            HSSFSheet example1 = (HSSFSheet)wb.GetSheet("Example 1");
            HSSFCell ex1cell1 = (HSSFCell)example1.GetRow(7).GetCell(0);
            Assert.AreEqual(3.0, fe.Evaluate(ex1cell1).NumberValue);
            HSSFCell ex1cell2 = (HSSFCell)example1.GetRow(8).GetCell(0);
            Assert.AreEqual(5.0, fe.Evaluate(ex1cell2).NumberValue);

            HSSFSheet example2 = (HSSFSheet)wb.GetSheet("Example 2");
            for (int rownum = 1; rownum <= 10; rownum++)
            {
                HSSFCell cell = (HSSFCell)example2.GetRow(rownum).GetCell(2);
                double cachedResult = cell.NumericCellValue; //cached formula result
                Assert.AreEqual(cachedResult, fe.Evaluate(cell).NumberValue);
            }

            HSSFSheet example3 = (HSSFSheet)wb.GetSheet("Example 3");
            for (int rownum = 1; rownum <= 10; rownum++)
            {
                HSSFCell cellD = (HSSFCell)example3.GetRow(rownum).GetCell(3);
                double cachedResultD = cellD.NumericCellValue; //cached formula result
                Assert.AreEqual(cachedResultD, fe.Evaluate(cellD).NumberValue, new CellReference(cellD).FormatAsString());

                HSSFCell cellE = (HSSFCell)example3.GetRow(rownum).GetCell(4);
                double cachedResultE = cellE.NumericCellValue; //cached formula result
                Assert.AreEqual(cachedResultE, fe.Evaluate(cellE).NumberValue, new CellReference(cellE).FormatAsString());

                HSSFCell cellF = (HSSFCell)example3.GetRow(rownum).GetCell(5);
                double cachedResultF = cellF.NumericCellValue; //cached formula result
                Assert.AreEqual(cachedResultF, fe.Evaluate(cellF).NumberValue, new CellReference(cellF).FormatAsString());
            }
        }
    }
}