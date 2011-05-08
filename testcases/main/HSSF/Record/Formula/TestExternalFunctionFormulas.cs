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

namespace TestCases.HSSF.Record.Formula
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.UserModel;
    /**
     * Tests for functions from external workbooks (e.g. YEARFRAC).
     * 
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestExternalFunctionFormulas
    {


        /**
         * tests <tt>NameXPtg.ToFormulaString(Workbook)</tt> and logic in Workbook below that   
         */
        [TestMethod]
        public void TestReadFormulaContainingExternalFunction()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("externalFunctionExample.xls");

            String expectedFormula = "YEARFRAC(B1,C1)";
            NPOI.SS.UserModel.Sheet sht = wb.GetSheetAt(0);
            String cellFormula = sht.GetRow(0).GetCell(0).CellFormula;
            Assert.AreEqual(expectedFormula, cellFormula);
        }

    }
}