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
    using NUnit.Framework;

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

            double accuracy = 1E-9;

            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            Assert.AreEqual(9.0, Evaluator.Evaluate(cell1).NumberValue, accuracy, "Rounds 10 to a nearest multiple of 3 (9)");

            Assert.AreEqual(-9.0, Evaluator.Evaluate(cell2).NumberValue, accuracy, "Rounds -10 to a nearest multiple of -3 (-9)");

            Assert.AreEqual(1.4, Evaluator.Evaluate(cell3).NumberValue, accuracy, "Rounds 1.3 to a nearest multiple of 0.2 (1.4)");

            Assert.AreEqual(ErrorEval.NUM_ERROR.ErrorCode, Evaluator.Evaluate(cell4).ErrorValue, "Returns an error, because -2 and 5 have different signs (#NUM!)");

            Assert.AreEqual(0.0, Evaluator.Evaluate(cell5).NumberValue, "Returns 0 because the multiple is 0");
        }
    }

}