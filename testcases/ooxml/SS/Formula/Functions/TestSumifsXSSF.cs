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

namespace TestCases.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF;
    using NUnit.Framework;
    [TestFixture]
    public class TestSumifsXSSF
    {
        [Test]
        public void TestBug60858()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("bug60858.xlsx");
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sheet = wb.GetSheetAt(0);
            ICell cell = sheet.GetRow(1).GetCell(5);
            fe.Evaluate(cell);
            Assert.AreEqual(0.0, cell.NumericCellValue, 0.0000000000000001);
        }
    }
}
