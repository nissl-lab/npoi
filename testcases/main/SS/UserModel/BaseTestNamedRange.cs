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

using NPOI.SS.UserModel;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.Util;
namespace TestCases.SS.UserModel
{

    /**
     * Tests of implementations of {@link NPOI.SS.usermodel.Name}.
     *
     * @author Yegor Kozlov
     */
    public abstract class BaseTestNamedRange
    {

        protected ITestDataProvider _testDataProvider;

        /**
         * @param testDataProvider an object that provides test data in HSSF / XSSF specific way
         */
        protected BaseTestNamedRange(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }
        [TestMethod]
        public void TestCreate()
        {
            // Create a new workbook
            Workbook wb = _testDataProvider.CreateWorkbook();
            Sheet sheet1 = wb.CreateSheet("Test1");
            Sheet sheet2 = wb.CreateSheet("Testing Named Ranges");

            Name name1 = wb.CreateName();
            name1.NameName = ("testOne");

            //Setting a duplicate name should throw ArgumentException
            Name name2 = wb.CreateName();
            try
            {
                name2.NameName = ("testOne");
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("The workbook already contains this name: testOne", e.Message);
            }
            //the check for duplicates is case-insensitive
            try
            {
                name2.NameName = ("TESTone");
                Assert.Fail("expected exception");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("The workbook already contains this name: TESTone", e.Message);
            }

            name2.NameName = ("testTwo");

            String ref1 = "Test1!$A$1:$B$1";
            name1.RefersToFormula = (ref1);
            Assert.AreEqual(ref1, name1.RefersToFormula);
            Assert.AreEqual("Test1", name1.SheetName);

            String ref2 = "'Testing Named Ranges'!$A$1:$B$1";
            name1.RefersToFormula = (ref2);
            Assert.AreEqual("'Testing Named Ranges'!$A$1:$B$1", name1.RefersToFormula);
            Assert.AreEqual("Testing Named Ranges", name1.SheetName);

            Assert.AreEqual(-1, name1.SheetIndex);
            name1.SheetIndex = (-1);
            Assert.AreEqual(-1, name1.SheetIndex);
            try
            {
                name1.SheetIndex = (2);
                Assert.Fail("should throw ArgumentException");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Sheet index (2) is out of range (0..1)", e.Message);
            }

            name1.SheetIndex = (1);
            Assert.AreEqual(1, name1.SheetIndex);

            //-1 means the name applies to the entire workbook
            name1.SheetIndex = (-1);
            Assert.AreEqual(-1, name1.SheetIndex);

            //names cannot be blank and must begin with a letter or underscore and not contain spaces
            String[] invalidNames = { "", "123", "1Name", "Named Range" };
            foreach (String name in invalidNames)
            {
                try
                {
                    name1.NameName = (name);
                    Assert.Fail("should have thrown exceptiuon due to invalid name: " + name);
                }
                catch (ArgumentException e)
                {
                    // expected during successful test
                    ;
                }
            }
        }
        [TestMethod]
        public void TestUnicodeNamedRange()
        {
            Workbook workBook = _testDataProvider.CreateWorkbook();
            workBook.CreateSheet("Test");
            Name name = workBook.CreateName();
            name.NameName = ("\u03B1");
            name.RefersToFormula = ("Test!$D$3:$E$8");


            Workbook workBook2 = _testDataProvider.WriteOutAndReadBack(workBook);
            Name name2 = workBook2.GetNameAt(0);

            Assert.AreEqual("\u03B1", name2.NameName);
            Assert.AreEqual("Test!$D$3:$E$8", name2.RefersToFormula);
        }
        [TestMethod]
        public void TestAddRemove()
        {
            Workbook wb = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, wb.NumberOfNames);
            Name name1 = wb.CreateName();
            name1.NameName = ("name1");
            Assert.AreEqual(1, wb.NumberOfNames);

            Name name2 = wb.CreateName();
            name2.NameName = ("name2");
            Assert.AreEqual(2, wb.NumberOfNames);

            Name name3 = wb.CreateName();
            name3.NameName = ("name3");
            Assert.AreEqual(3, wb.NumberOfNames);

            wb.RemoveName("name2");
            Assert.AreEqual(2, wb.NumberOfNames);

            wb.RemoveName(0);
            Assert.AreEqual(1, wb.NumberOfNames);
        }
        [TestMethod]
        public void TestScope()
        {
            Workbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet();
            wb.CreateSheet();

            Name name;

            name = wb.CreateName();
            name.NameName = ("aaa");
            name = wb.CreateName();
            try
            {
                name.NameName = ("aaa");
                Assert.Fail("Expected exception");
            }
            catch (Exception e)
            {
                Assert.AreEqual("The workbook already contains this name: aaa", e.Message);
            }

            name = wb.CreateName();
            name.SheetIndex = (0);
            name.NameName = ("aaa");
            name = wb.CreateName();
            name.SheetIndex = (0);
            try
            {
                name.NameName = ("aaa");
                Assert.Fail("Expected exception");
            }
            catch (Exception e)
            {
                Assert.AreEqual("The sheet already contains this name: aaa", e.Message);
            }

            name = wb.CreateName();
            name.SheetIndex = (1);
            name.NameName = ("aaa");
            name = wb.CreateName();
            name.SheetIndex = (1);
            try
            {
                name.NameName = ("aaa");
                Assert.Fail("Expected exception");
            }
            catch (Exception e)
            {
                Assert.AreEqual("The sheet already contains this name: aaa", e.Message);
            }

            int cnt = 0;
            for (int i = 0; i < wb.NumberOfNames; i++)
            {
                if ("aaa".Equals(wb.GetNameAt(i).NameName)) cnt++;
            }
            Assert.AreEqual(3, cnt);
        }

        /**
         * Test case provided by czhang@cambian.com (Chun Zhang)
         * <p>
         * Addresses Bug <a href="http://issues.apache.org/bugzilla/Show_bug.cgi?id=13775" tarGet="_bug">#13775</a>
         */
        [TestMethod]
        public void TestMultiNamedRange()
        {

            // Create a new workbook
            Workbook wb = _testDataProvider.CreateWorkbook();

            // Create a worksheet 'sheet1' in the new workbook
            wb.CreateSheet();
            wb.SetSheetName(0, "sheet1");

            // Create another worksheet 'sheet2' in the new workbook
            wb.CreateSheet();
            wb.SetSheetName(1, "sheet2");

            // Create a new named range for worksheet 'sheet1'
            Name namedRange1 = wb.CreateName();

            // Set the name for the named range for worksheet 'sheet1'
            namedRange1.NameName = ("RangeTest1");

            // Set the reference for the named range for worksheet 'sheet1'
            namedRange1.RefersToFormula = ("sheet1" + "!$A$1:$L$41");

            // Create a new named range for worksheet 'sheet2'
            Name namedRange2 = wb.CreateName();

            // Set the name for the named range for worksheet 'sheet2'
            namedRange2.NameName = ("RangeTest2");

            // Set the reference for the named range for worksheet 'sheet2'
            namedRange2.RefersToFormula = ("sheet2" + "!$A$1:$O$21");

            // Write the workbook to a file
            // Read the Excel file and verify its content
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            Name nm1 = wb.GetNameAt(wb.GetNameIndex("RangeTest1"));
            Assert.IsTrue("RangeTest1".Equals(nm1.NameName), "Name is " + nm1.NameName);
            Assert.IsTrue((wb.GetSheetName(0) + "!$A$1:$L$41").Equals(nm1.RefersToFormula), "Reference is " + nm1.RefersToFormula);

            Name nm2 = wb.GetNameAt(wb.GetNameIndex("RangeTest2"));
            Assert.IsTrue("RangeTest2".Equals(nm2.NameName), "Name is " + nm2.NameName);
            Assert.IsTrue((wb.GetSheetName(1) + "!$A$1:$O$21").Equals(nm2.RefersToFormula), "Reference is " + nm2.RefersToFormula);
        }

        /**
         * Test to see if the print areas can be retrieved/Created in memory
         */
        [TestMethod]
        public void TestSinglePrintArea()
        {
            Workbook workbook = _testDataProvider.CreateWorkbook();
            workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);

            String reference = "$A$1:$B$1";
            workbook.SetPrintArea(0, reference);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!$A$1:$B$1", retrievedPrintArea);
        }

        /**
         * For Convenience, don't force sheet names to be used
         */
        [TestMethod]
        public void TestSinglePrintAreaWOSheet()
        {
            Workbook workbook = _testDataProvider.CreateWorkbook();
            workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);

            String reference = "$A$1:$B$1";
            workbook.SetPrintArea(0, reference);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!" + reference, retrievedPrintArea);
        }

        /**
         * Test to see if the print area made it to the file
         */
        [TestMethod]
        public void TestPrintAreaFile()
        {
            Workbook workbook = _testDataProvider.CreateWorkbook();
            workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);


            String reference = "$A$1:$B$1";
            workbook.SetPrintArea(0, reference);

            workbook = _testDataProvider.WriteOutAndReadBack(workbook);

            String retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.IsNotNull(retrievedPrintArea,"Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!$A$1:$B$1", retrievedPrintArea, "References Match");
        }

        /**
         * Test to see if multiple print areas made it to the file
         */
        [TestMethod]
        public void TestMultiplePrintAreaFile()
        {
            Workbook workbook = _testDataProvider.CreateWorkbook();

            workbook.CreateSheet("Sheet1");
            workbook.CreateSheet("Sheet2");
            workbook.CreateSheet("Sheet3");
            String reference1 = "$A$1:$B$1";
            String reference2 = "$B$2:$D$5";
            String reference3 = "$D$2:$F$5";

            workbook.SetPrintArea(0, reference1);
            workbook.SetPrintArea(1, reference2);
            workbook.SetPrintArea(2, reference3);

            //Check Created print areas
            String retrievedPrintArea;

            retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 1)");
            Assert.AreEqual("Sheet1!" + reference1, retrievedPrintArea);

            retrievedPrintArea = workbook.GetPrintArea(1);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 2)");
            Assert.AreEqual("Sheet2!" + reference2, retrievedPrintArea);

            retrievedPrintArea = workbook.GetPrintArea(2);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 3)");
            Assert.AreEqual("Sheet3!" + reference3, retrievedPrintArea);

            // Check print areas after re-reading workbook
            workbook = _testDataProvider.WriteOutAndReadBack(workbook);

            retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 1)");
            Assert.AreEqual("Sheet1!" + reference1, retrievedPrintArea);

            retrievedPrintArea = workbook.GetPrintArea(1);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 2)");
            Assert.AreEqual("Sheet2!" + reference2, retrievedPrintArea);

            retrievedPrintArea = workbook.GetPrintArea(2);
            Assert.IsNotNull(retrievedPrintArea, "Print Area Not Found (Sheet 3)");
            Assert.AreEqual("Sheet3!" + reference3, retrievedPrintArea);
        }

        /**
         * Tests the Setting of print areas with coordinates (Row/Column designations)
         *
         */
        [TestMethod]
        public void TestPrintAreaCoords()
        {
            Workbook workbook = _testDataProvider.CreateWorkbook();
            workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);

            workbook.SetPrintArea(0, 0, 1, 0, 0);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!$A$1:$B$1", retrievedPrintArea);
        }


        /**
         * Tests the parsing of union area expressions, and re-display in the presence of sheet names
         * with special characters.
         */
        [TestMethod]
        public void TestPrintAreaUnion()
        {
            Workbook workbook = _testDataProvider.CreateWorkbook();
            workbook.CreateSheet("Test Print Area");

            String reference = "$A$1:$B$1,$D$1:$F$2";
            workbook.SetPrintArea(0, reference);
            String retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'Test Print Area'!$A$1:$B$1,'Test Print Area'!$D$1:$F$2", retrievedPrintArea);
        }

        /**
         * Verifies an existing print area is deleted
         *
         */
        [TestMethod]
        public void TestPrintAreaRemove()
        {
            Workbook workbook = _testDataProvider.CreateWorkbook();
            workbook.CreateSheet("Test Print Area");
            workbook.GetSheetName(0);

            workbook.SetPrintArea(0, 0, 1, 0, 0);

            String retrievedPrintArea = workbook.GetPrintArea(0);

            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");

            workbook.RemovePrintArea(0);
            Assert.IsNull(workbook.GetPrintArea(0), "PrintArea was not Removed");
        }

        /**
         * Test that multiple named ranges can be Added written and read
         */
        [TestMethod]
        public void TestMultipleNamedWrite()
        {
            Workbook wb = _testDataProvider.CreateWorkbook();


            wb.CreateSheet("testSheet1");
            String sheetName = wb.GetSheetName(0);

            Assert.AreEqual("testSheet1", sheetName);

            //Creating new Named Range
            Name newNamedRange = wb.CreateName();

            newNamedRange.NameName = ("RangeTest");
            newNamedRange.RefersToFormula = (sheetName + "!$D$4:$E$8");

            //Creating another new Named Range
            Name newNamedRange2 = wb.CreateName();

            newNamedRange2.NameName = ("AnotherTest");
            newNamedRange2.RefersToFormula = (sheetName + "!$F$1:$G$6");

            wb.GetNameAt(0);

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            Name nm = wb.GetNameAt(wb.GetNameIndex("RangeTest"));
            Assert.IsTrue("RangeTest".Equals(nm.NameName), "Name is " + nm.NameName);
            Assert.IsTrue((wb.GetSheetName(0) + "!$D$4:$E$8").Equals(nm.RefersToFormula), "Reference is " + nm.RefersToFormula);

            nm = wb.GetNameAt(wb.GetNameIndex("AnotherTest"));
            Assert.IsTrue("AnotherTest".Equals(nm.NameName), "Name is " + nm.NameName);
            Assert.IsTrue(newNamedRange2.RefersToFormula.Equals(nm.RefersToFormula), "Reference is " + nm.RefersToFormula);
        }
        /**
         * Verifies correct functioning for "single cell named range" (aka "named cell")
         */
        [TestMethod]
        public void TestNamedCell_1()
        {

            // Setup for this testcase
            String sheetName = "Test Named Cell";
            String cellName = "named_cell";
            String cellValue = "TEST Value";
            Workbook wb = _testDataProvider.CreateWorkbook();
            Sheet sheet = wb.CreateSheet(sheetName);
            CreationHelper factory = wb.GetCreationHelper();
            sheet.CreateRow(0).CreateCell(0).SetCellValue(factory.CreateRichTextString(cellValue));

            // create named range for a single cell using areareference
            Name namedCell = wb.CreateName();
            namedCell.NameName = (cellName);
            String reference = "'" + sheetName + "'" + "!A1:A1";
            namedCell.RefersToFormula = (reference);

            // retrieve the newly Created named range
            int namedCellIdx = wb.GetNameIndex(cellName);
            Name aNamedCell = wb.GetNameAt(namedCellIdx);
            Assert.IsNotNull(aNamedCell);

            // retrieve the cell at the named range and test its contents
            AreaReference aref = new AreaReference(aNamedCell.RefersToFormula);
            Assert.IsTrue(aref.IsSingleCell, "Should be exactly 1 cell in the named cell :'" + cellName + "'");

            CellReference cref = aref.FirstCell;
            Assert.IsNotNull(cref);
            Sheet s = wb.GetSheet(cref.SheetName);
            Assert.IsNotNull(s);
            Row r = sheet.GetRow(cref.Row);
            Cell c = r.GetCell(cref.Col);
            String contents = c.RichStringCellValue.String;
            Assert.AreEqual(contents, cellValue, "Contents of cell retrieved by its named reference");
        }

        /**
         * Verifies correct functioning for "single cell named range" (aka "named cell")
         */
        [TestMethod]
        public void TestNamedCell_2()
        {

            // Setup for this testcase
            String sname = "TestSheet", cname = "TestName", cvalue = "TestVal";
            Workbook wb = _testDataProvider.CreateWorkbook();
            CreationHelper factory = wb.GetCreationHelper();
            Sheet sheet = wb.CreateSheet(sname);
            sheet.CreateRow(0).CreateCell(0).SetCellValue(factory.CreateRichTextString(cvalue));

            // create named range for a single cell using cellreference
            Name namedCell = wb.CreateName();
            namedCell.NameName = (cname);
            String reference = sname + "!A1";
            namedCell.RefersToFormula = (reference);

            // retrieve the newly Created named range
            int namedCellIdx = wb.GetNameIndex(cname);
            Name aNamedCell = wb.GetNameAt(namedCellIdx);
            Assert.IsNotNull(aNamedCell);

            // retrieve the cell at the named range and test its contents
            CellReference cref = new CellReference(aNamedCell.RefersToFormula);
            Assert.IsNotNull(cref);
            Sheet s = wb.GetSheet(cref.SheetName);
            Row r = sheet.GetRow(cref.Row);
            Cell c = r.GetCell(cref.Col);
            String contents = c.RichStringCellValue.String;
            Assert.AreEqual(contents, cvalue, "Contents of cell retrieved by its named reference");
        }


        /**
         * Bugzilla attachment 23444 (from bug 46973) has a NAME record with the following encoding: 
         * <pre>
         * 00000000 | 18 00 17 00 00 00 00 08 00 00 00 00 00 00 00 00 | ................
         * 00000010 | 00 00 00 55 50 53 53 74 61 74 65                | ...UPSState
         * </pre>     
         * 
         * This caused trouble for anything that requires {@link HSSFName#RefersToFormula}
         * It is easy enough to re-create the the same data (by not Setting the formula). Excel
         * seems to gracefully remove this unInitialized name record.  It would be nice if POI
         * could do the same, but that would involve adjusting subsequent name indexes across 
         * all formulas. <p/>
         * 
         * For the moment, POI has been made to behave more sensibly with unInitialised name 
         * records.
         */
        [TestMethod]
        public void TestUnInitialisedNameRefersToFormula_bug46973()
        {
            Workbook wb = _testDataProvider.CreateWorkbook();
            Name n = wb.CreateName();
            n.NameName = ("UPSState");
            String formula;
            try
            {
                formula = n.RefersToFormula;
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("ptgs must not be null"))
                {
                    throw new AssertFailedException("Identified bug 46973");
                }
                throw e;
            }
            Assert.IsNull(formula);
            Assert.IsFalse(n.IsDeleted); // according to exact defInition of IsDeleted
        }
        [TestMethod]
        public void TestDeletedCell()
        {
            Workbook wb = _testDataProvider.CreateWorkbook();
            Name n = wb.CreateName();
            n.NameName = ("MyName");
            // contrived example to expose bug:
            n.RefersToFormula = ("if(A1,\"#REF!\", \"\")");

            if (n.IsDeleted)
            {
                throw new AssertFailedException("Identified bug in recoginising formulas referring to deleted cells");
            }

        }
        [TestMethod]
        public void TestFunctionNames()
        {
            Workbook wb = _testDataProvider.CreateWorkbook();
            Name n = wb.CreateName();
            Assert.IsFalse(n.IsFunctionName);

            n.IsFunctionName =false;
            Assert.IsFalse(n.IsFunctionName);

            n.IsFunctionName = true;
            Assert.IsTrue(n.IsFunctionName);

            n.IsFunctionName = false; 
            Assert.IsFalse(n.IsFunctionName);
        }
    }
}



