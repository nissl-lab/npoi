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

namespace TestCases.SS.Formula.PTG
{
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;
    /**
     * Tests for Ref3DPtg
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRef3DPtg : AbstractPtgTestCase
    {
        [Test]
        public void TestToFormulaString()
        {

            Ref3DPtg target = new Ref3DPtg("A1", (short)0);

            HSSFWorkbook wb = CreateWorkbookWithSheet("my sheet");
            HSSFEvaluationWorkbook book = HSSFEvaluationWorkbook.Create(wb);
            Assert.AreEqual("'my sheet'!A1", target.ToFormulaString(book));

            wb.SetSheetName(0, "ProfitAndLoss");
            Assert.AreEqual("ProfitAndLoss!A1", target.ToFormulaString(book));

            wb.SetSheetName(0, "profit+loss");
            Assert.AreEqual("'profit+loss'!A1", target.ToFormulaString(book));
        }
    }

}