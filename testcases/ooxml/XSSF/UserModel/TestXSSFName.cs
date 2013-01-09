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

using TestCases.SS.UserModel;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
namespace NPOI.XSSF.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFName : BaseTestNamedRange
    {

        public TestXSSFName()
            : base(XSSFITestDataProvider.instance)
        {

        }

        //TODO combine TestRepeatingRowsAndColums() for HSSF and XSSF
        [Test]
        public void TestRepeatingRowsAndColums()
        {
            // First Test that Setting RR&C for same sheet more than once only Creates a
            // single  Print_Titles built-in record
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)wb.CreateSheet("First Sheet");
            sheet1.RepeatingRows = (null);
            sheet1.RepeatingColumns = (null);



            // Set repeating rows and columns twice for the first sheet
            for (int i = 0; i < 2; i++)
            {
                sheet1.RepeatingRows = (CellRangeAddress.ValueOf("1:4"));
                sheet1.RepeatingColumns = (CellRangeAddress.ValueOf("A:A"));
                //sheet.CreateFreezePane(0, 3);
            }
            Assert.AreEqual(1, wb.NumberOfNames);
            IName nr1 = wb.GetNameAt(0);

            Assert.AreEqual(XSSFName.BUILTIN_PRINT_TITLE, nr1.NameName);
            Assert.AreEqual("'First Sheet'!$A:$A,'First Sheet'!$1:$4", nr1.RefersToFormula);

            //remove the columns part
            sheet1.RepeatingColumns = (null);
            Assert.AreEqual("'First Sheet'!$1:$4", nr1.RefersToFormula);

            //revert
            sheet1.RepeatingColumns = (CellRangeAddress.ValueOf("A:A"));

            //remove the rows part
            sheet1.RepeatingRows=(null);
            Assert.AreEqual("'First Sheet'!$A:$A", nr1.RefersToFormula);

            //revert
            sheet1.RepeatingRows = (CellRangeAddress.ValueOf("1:4"));

            // Save and re-open
            IWorkbook nwb = XSSFTestDataSamples.WriteOutAndReadBack(wb);

            Assert.AreEqual(1, nwb.NumberOfNames);
            nr1 = nwb.GetNameAt(0);

            Assert.AreEqual(XSSFName.BUILTIN_PRINT_TITLE, nr1.NameName);
            Assert.AreEqual("'First Sheet'!$A:$A,'First Sheet'!$1:$4", nr1.RefersToFormula);

            // check that Setting RR&C on a second sheet causes a new Print_Titles built-in
            // name to be Created
            XSSFSheet sheet2 = (XSSFSheet)nwb.CreateSheet("SecondSheet");
            sheet2.RepeatingRows = (CellRangeAddress.ValueOf("1:1"));
            sheet2.RepeatingColumns = (CellRangeAddress.ValueOf("B:C"));

            Assert.AreEqual(2, nwb.NumberOfNames);
            IName nr2 = nwb.GetNameAt(1);

            Assert.AreEqual(XSSFName.BUILTIN_PRINT_TITLE, nr2.NameName);
            Assert.AreEqual("SecondSheet!$B:$C,SecondSheet!$1:$1", nr2.RefersToFormula);

            sheet2.RepeatingRows = (null);
            sheet2.RepeatingColumns = (null);
        }
    }
}

