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
namespace TestCases.SS.Formula.Atp
{

    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    /**
     * Testcase for 'Analysis Toolpak' function MROUND()
     * 
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestMRound
    {

        /**
    =MROUND(10, 3) 	Rounds 10 to a nearest multiple of 3 (9)
    =MROUND(-10, -3) 	Rounds -10 to a nearest multiple of -3 (-9)
    =MROUND(1.3, 0.2) 	Rounds 1.3 to a nearest multiple of 0.2 (1.4)
    =MROUND(5, -2) 	Returns an error, because -2 and 5 have different signs (#NUM!)     *
         */
        [Test]
        public void TestEvaluate()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US"); 
            
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sh = wb.CreateSheet();
            ICell cell1 = sh.CreateRow(0).CreateCell(0);
            cell1.CellFormula = ("MROUND(10, 3)");
            ICell cell2 = sh.CreateRow(0).CreateCell(0);
            cell2.CellFormula = ("MROUND(-10, -3)");
            ICell cell3 = sh.CreateRow(0).CreateCell(0);
            cell3.CellFormula = ("MROUND(1.3, 0.2)");
            ICell cell4 = sh.CreateRow(0).CreateCell(0);
            cell4.CellFormula = ("MROUND(5, -2)");
            ICell cell5 = sh.CreateRow(0).CreateCell(0);
            cell5.CellFormula = ("MROUND(5, 0)");
            ICell cell6 = sh.CreateRow(0).CreateCell(0);
            cell6.CellFormula = ("MROUND(0.79*7.5, 0.05)");

            double accuracy = 1E-9;

            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            ClassicAssert.AreEqual(9.0, Evaluator.Evaluate(cell1).NumberValue, accuracy, "Rounds 10 to a nearest multiple of 3 (9)");

            ClassicAssert.AreEqual(-9.0, Evaluator.Evaluate(cell2).NumberValue, accuracy, "Rounds -10 to a nearest multiple of -3 (-9)");

            ClassicAssert.AreEqual(1.4, Evaluator.Evaluate(cell3).NumberValue, accuracy, "Rounds 1.3 to a nearest multiple of 0.2 (1.4)");

            ClassicAssert.AreEqual(ErrorEval.NUM_ERROR.ErrorCode, Evaluator.Evaluate(cell4).ErrorValue, "Returns an error, because -2 and 5 have different signs (#NUM!)");

            ClassicAssert.AreEqual(0.0, Evaluator.Evaluate(cell5).NumberValue, "Returns 0 because the multiple is 0");

            ClassicAssert.AreEqual(5.95, Evaluator.Evaluate(cell6).NumberValue, 0, "Rounds 5.925 to a nearest multiple of 0.05 (5.95)");
        }

        [Test]
        public void TestRoundsOnExcelsFifteenDigitView()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            Confirm(evaluator, cell, "MROUND(0.7+0.1, 1.6)", 1.6);
            Confirm(evaluator, cell, "MROUND(-0.7-0.1, -1.6)", -1.6);
            Confirm(evaluator, cell, "MROUND(880000000*0.00849/3, 1)", 2490400);
            Confirm(evaluator, cell, "MROUND(2.5, 1)", 3);
            Confirm(evaluator, cell, "MROUND(-2.5, -1)", -3);
            Confirm(evaluator, cell, "MROUND(2.4999999999999996, 1)", 3);
            Confirm(evaluator, cell, "MROUND(2.49999999999999, 1)", 2);
            Confirm(evaluator, cell, "MROUND(0.45, 0.1*3)", 0.6);
            Confirm(evaluator, cell, "MROUND(1.2, 0.7+0.1)", 1.6);
        }

        [Test]
        public void TestOutsideDecimalRange()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            Confirm(evaluator, cell, "MROUND(1E+300, 3E+299)", 9E+299);
            Confirm(evaluator, cell, "MROUND(-1E+30, -3)", -1E+30);
            Confirm(evaluator, cell, "MROUND(1.5E-30, 1E-30)", 2E-30);
            Confirm(evaluator, cell, "MROUND(4.5E-21, 1E-21*3)", 6E-21);
            Confirm(evaluator, cell, "MROUND(0.7+0.1, 1.6E-30)", 0.8);
            Confirm(evaluator, cell, "MROUND(1E-20, 1)", 0);
        }

        private static void Confirm(IFormulaEvaluator evaluator, ICell cell, string formula, double expected)
        {
            cell.SetCellFormula(formula);
            evaluator.NotifyUpdateCell(cell);
            CellValue value = evaluator.Evaluate(cell);
            ClassicAssert.AreEqual(CellType.Numeric, value.CellType, formula);
            ClassicAssert.AreEqual(expected, value.NumberValue, 0, formula);
        }
    }

}