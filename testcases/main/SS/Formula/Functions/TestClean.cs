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
namespace TestCases.SS.Formula.Functions
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using System;
    using NUnit.Framework;
    [TestFixture]
    public class TestClean
    {
        [Test]
        public void TestClean1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            String[] asserts = {
            "aniket\u0007\u0017\u0019", "aniket",
            "\u0011aniket\u0007\u0017\u0010", "aniket",
            "\u0011aniket\u0007\u0017\u007F", "aniket\u007F",
            "\u2116aniket\u2211\uFB5E\u2039", "\u2116aniket\u2211\uFB5E\u2039",
        };

            for (int i = 0; i < asserts.Length; i += 2)
            {
                String formulaText = "CLEAN(\"" + asserts[i] + "\")";
                ConfirmResult(fe, cell, formulaText, asserts[i + 1]);
            }

            asserts = new String[] {
            "CHAR(7)&\"text\"&CHAR(7)", "text",
            "CHAR(7)&\"text\"&CHAR(17)", "text",
            "CHAR(181)&\"text\"&CHAR(190)", "\u00B5text\u00BE",
            "\"text\"&CHAR(160)&\"'\"", "text\u00A0'",
        };
            for (int i = 0; i < asserts.Length; i += 2)
            {
                String formulaText = "CLEAN(" + asserts[i] + ")";
                ConfirmResult(fe, cell, formulaText, asserts[i + 1]);
            }
        }

        private static void ConfirmResult(HSSFFormulaEvaluator fe, ICell cell, String formulaText,
                                          String expectedResult)
        {
            cell.CellFormula = (formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(result.CellType, CellType.String);
            Assert.AreEqual(expectedResult, result.StringValue);
        }
    }

}