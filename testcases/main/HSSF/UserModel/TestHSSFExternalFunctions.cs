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

namespace TestCases.HSSF.UserModel
{
    using System;

    using NPOI.HSSF;
    using NPOI.SS.Formula;
    using TestCases.SS.Formula;
    using TestCases.HSSF;
    using NUnit.Framework;
    using NPOI.SS.UserModel;

    /**
     * Tests Setting and Evaluating user-defined functions in HSSF
     */
    [TestFixture]
    public class TestHSSFExternalFunctions //: BaseTestExternalFunctions
    {
        //public TestHSSFExternalFunctions()
        //    : base(HSSFITestDataProvider.Instance, "atp.xls")
        //{
            
        //}
        /* This test is a copy of BaseTestExternalFunctions.BaseTestInvokeATP(String testFile)
         * If we made this test class derived from BaseTestExternalFunctions, 
         * the test BaseTestExternalFunctions.TestExternalFunctions would run twice,
         * the fist passed and the second failed.
         * 
         */
        [Test]
        public void TestATP()
        {
            //baseTestInvokeATP("atp.xls");
            IWorkbook wb = HSSFITestDataProvider.Instance.OpenSampleWorkbook("atp.xls");
            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sh = wb.GetSheetAt(0);
            // these two are not imlemented in r
            Assert.AreEqual("DELTA(1.3,1.5)", sh.GetRow(0).GetCell(1).CellFormula);
            Assert.AreEqual("COMPLEX(2,4)", sh.GetRow(1).GetCell(1).CellFormula);

            ICell cell2 = sh.GetRow(2).GetCell(1);
            Assert.AreEqual("ISODD(2)", cell2.CellFormula);
            Assert.AreEqual(false, Evaluator.Evaluate(cell2).BooleanValue);
            Assert.AreEqual(CellType.Boolean, Evaluator.EvaluateFormulaCell(cell2));

            ICell cell3 = sh.GetRow(3).GetCell(1);
            Assert.AreEqual("ISEVEN(2)", cell3.CellFormula);
            Assert.AreEqual(true, Evaluator.Evaluate(cell3).BooleanValue);
            Assert.AreEqual(CellType.Boolean, Evaluator.EvaluateFormulaCell(cell3));
        }

    }

}