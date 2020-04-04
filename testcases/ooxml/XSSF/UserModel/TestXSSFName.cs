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
using NPOI.XSSF;
using NPOI.XSSF.UserModel;

namespace TestCases.XSSF.UserModel
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
            IName nr1 = wb.GetName(XSSFName.BUILTIN_PRINT_TITLE);

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
            nr1 = nwb.GetName(XSSFName.BUILTIN_PRINT_TITLE);

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

        [Test]
        public void TestSetNameName()
        {
            // Test that renaming named ranges doesn't break our new named range map
            XSSFWorkbook wb = new XSSFWorkbook();
            wb.CreateSheet("First Sheet");
            // Two named ranges called "name1", one scoped to sheet1 and one globally
            XSSFName nameSheet1 = wb.CreateName() as XSSFName;
            nameSheet1.NameName = "name1";
            nameSheet1.RefersToFormula = "'First Sheet'!$A$1";
            nameSheet1.SheetIndex = 0;
            XSSFName nameGlobal = wb.CreateName() as XSSFName;
            nameGlobal.NameName = "name1";
            nameGlobal.RefersToFormula = "'First Sheet'!$B$1";
            // Rename sheet-scoped name to "name2", check everything is updated properly
            // and that the other name is unaffected
            nameSheet1.NameName = "name2";
            Assert.AreEqual(1, wb.GetNames("name1").Count);
            Assert.AreEqual(1, wb.GetNames("name2").Count);
            Assert.AreEqual(nameGlobal, wb.GetName("name1"));
            Assert.AreEqual(nameSheet1, wb.GetName("name2"));
            // Rename the other name to "name" and check everything again
            nameGlobal.NameName = "name2";
            Assert.AreEqual(0, wb.GetNames("name1").Count);
            Assert.AreEqual(2, wb.GetNames("name2").Count);
            Assert.IsTrue(wb.GetNames("name2").Contains(nameGlobal));
            Assert.IsTrue(wb.GetNames("name2").Contains(nameSheet1));
            wb.Close();
        }

    }
}

