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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using System.Text;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;

    [TestFixture]
    public class TestProper
    {
        private ICell cell11;
        private IFormulaEvaluator Evaluator;

        [Test]
        public void TestValidHSSF()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Evaluator = new HSSFFormulaEvaluator(wb);

            Confirm(wb);
        }

        [Test]
        public void TestValidXSSF()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            Evaluator = new XSSFFormulaEvaluator(wb);

            Confirm(wb);
        }

        private void Confirm(IWorkbook wb)
        {
            ISheet sheet = wb.CreateSheet("new sheet");
            cell11 = sheet.CreateRow(0).CreateCell(0);
            cell11.SetCellType(CellType.Formula);

            Confirm("PROPER(\"hi there\")", "Hi There");
            Confirm("PROPER(\"what's up\")", "What'S Up");
            Confirm("PROPER(\"I DON'T TH!NK SO!\")", "I Don'T Th!Nk So!");
            Confirm("PROPER(\"dr\u00dcb\u00f6'\u00e4 \u00e9lo\u015f|\u00eb\u00e8 \")", "Dr\u00fcb\u00f6'\u00c4 \u00c9lo\u015f|\u00cb\u00e8 ");
            Confirm("PROPER(\"hi123 the123re\")", "Hi123 The123Re");
            Confirm("PROPER(\"-\")", "-");
            Confirm("PROPER(\"!\u00a7$\")", "!\u00a7$");
            Confirm("PROPER(\"/&%\")", "/&%");

            // also test longer string
            StringBuilder builder = new StringBuilder("A");
            StringBuilder expected = new StringBuilder("A");
            for (int i = 1; i < 254; i++)
            {
                builder.Append((char)(65 + (i % 26)));
                expected.Append((char)(97 + (i % 26)));
            }
            Confirm("PROPER(\"" + builder.ToString() + "\")", expected.ToString());
        }

        private void Confirm(String formulaText, String expectedResult)
        {
            cell11.CellFormula = (/*setter*/formulaText);
            Evaluator.ClearAllCachedResultValues();
            CellValue cv = Evaluator.Evaluate(cell11);
            if (cv.CellType != CellType.String)
            {
                Assert.Fail("Wrong result type: " + cv.FormatAsString());
            }
            String actualValue = cv.StringValue;
            Assert.AreEqual(expectedResult, actualValue);
        }
    }

}