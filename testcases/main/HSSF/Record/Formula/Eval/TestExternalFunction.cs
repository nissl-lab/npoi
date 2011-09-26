/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record.Formula.Eval
{
    using System;
    using NPOI.HSSF.UserModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record.Formula.Eval;
    using NPOI.SS.UserModel;

    /**
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestExternalFunction
    {

        /**
         * Checks that an external function can Get Invoked from the formula evaluator. 
         */
        [TestMethod]
        public void TestInvoke()
        {
            HSSFWorkbook wb;
            NPOI.SS.UserModel.ISheet sheet ;
            ICell cell;
            if (false)
            {
                // TODO - this code won't work until we can create user-defined functions directly with POI
                wb = new HSSFWorkbook();
                sheet = wb.CreateSheet();
                wb.SetSheetName(0, "Sheet1");
                NPOI.SS.UserModel.IName hssfName = wb.CreateName();
                hssfName.NameName = ("myFunc");

            }
            else
            {
                // This sample spreadsheet already has a VB function called 'myFunc'
                wb = HSSFTestDataSamples.OpenSampleWorkbook("testNames.xls");
                sheet = wb.GetSheetAt(0);
                IRow row = sheet.CreateRow(0);
                cell = row.CreateCell(1);
            }

            cell.CellFormula = ("myFunc()");
            String actualFormula = cell.CellFormula;
            Assert.AreEqual("myFunc()", actualFormula);

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(sheet, wb);
            NPOI.SS.UserModel.CellValue evalResult = fe.Evaluate(cell);

            // Check the return value from ExternalFunction.evaluate()
            // TODO - make this test assert something more interesting as soon as ExternalFunction works a bit better
            Assert.AreEqual(NPOI.SS.UserModel.CellType.ERROR, evalResult.CellType);
            Assert.AreEqual(ErrorEval.FUNCTION_NOT_IMPLEMENTED.ErrorCode, evalResult.ErrorValue);
        }
    }
}