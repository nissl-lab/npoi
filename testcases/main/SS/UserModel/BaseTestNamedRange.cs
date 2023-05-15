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

namespace TestCases.SS.UserModel
{
    using System;
    using NUnit.Framework;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using System.Collections.Generic;
    using NPOI.Util;
    using NPOI.HSSF.Util;

    /**
     * Tests of implementations of {@link NPOI.ss.usermodel.Name}.
     *
     * @author Yegor Kozlov
     */
    public abstract class BaseTestNamedRange
    {

        private ITestDataProvider _testDataProvider;
        //public BaseTestNamedRange()
        //    : this(TestCases.HSSF.HSSFITestDataProvider.Instance)
        //{ }
        protected BaseTestNamedRange(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }
        [Test]
        public void TestCreate()
        {
            // Create a new workbook
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet1 = wb.CreateSheet("Test1");
            ISheet sheet2 = wb.CreateSheet("Testing Named Ranges");

            IName name1 = wb.CreateName();
            name1.NameName = ("testOne");

            //setting a duplicate name should throw ArgumentException
            IName name2 = wb.CreateName();
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
                catch (ArgumentException)
                {
                    // expected during successful Test
                }
            }
        }
        [Test]
        public void TestUnicodeNamedRange()
        {
            IWorkbook workBook = _testDataProvider.CreateWorkbook();
            workBook.CreateSheet("Test");
            IName name = workBook.CreateName();
            name.NameName = ("\u03B1");
            name.RefersToFormula = ("Test!$D$3:$E$8");


            IWorkbook workBook2 = _testDataProvider.WriteOutAndReadBack(workBook);
            IName name2 = workBook2.GetNameAt(0);

            Assert.AreEqual("\u03B1", name2.NameName);
            Assert.AreEqual("Test!$D$3:$E$8", name2.RefersToFormula);
        }
        [Test]
        public void TestAddRemove()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            Assert.AreEqual(0, wb.NumberOfNames);
            IName name1 = wb.CreateName();
            name1.NameName = ("name1");
            Assert.AreEqual(1, wb.NumberOfNames);

            IName name2 = wb.CreateName();
            name2.NameName = ("name2");
            Assert.AreEqual(2, wb.NumberOfNames);

            IName name3 = wb.CreateName();
            name3.NameName = ("name3");
            Assert.AreEqual(3, wb.NumberOfNames);

            wb.RemoveName("name2");
            Assert.AreEqual(2, wb.NumberOfNames);

            wb.RemoveName(0);
            Assert.AreEqual(1, wb.NumberOfNames);
        }
        [Test]
        public void TestScope()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            wb.CreateSheet();
            wb.CreateSheet();

            IName name;

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

            Assert.AreEqual(3, wb.GetNames("aaa").Count);
        }

        /**
         * Test case provided by czhang@cambian.com (Chun Zhang)
         * <p>
         * Addresses Bug <a href="http://issues.apache.org/bugzilla/Show_bug.cgi?id=13775" target="_bug">#13775</a>
         */
        [Test]
        public void TestMultiNamedRange()
        {

            // Create a new workbook
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            // Create a worksheet 'sheet1' in the new workbook
            wb.CreateSheet();
            wb.SetSheetName(0, "sheet1");

            // Create another worksheet 'sheet2' in the new workbook
            wb.CreateSheet();
            wb.SetSheetName(1, "sheet2");

            // Create a new named range for worksheet 'sheet1'
            IName namedRange1 = wb.CreateName();

            // Set the name for the named range for worksheet 'sheet1'
            namedRange1.NameName = ("RangeTest1");

            // Set the reference for the named range for worksheet 'sheet1'
            namedRange1.RefersToFormula = ("sheet1" + "!$A$1:$L$41");

            // Create a new named range for worksheet 'sheet2'
            IName namedRange2 = wb.CreateName();

            // Set the name for the named range for worksheet 'sheet2'
            namedRange2.NameName = ("RangeTest2");

            // Set the reference for the named range for worksheet 'sheet2'
            namedRange2.RefersToFormula = ("sheet2" + "!$A$1:$O$21");

            // Write the workbook to a file
            // Read the Excel file and verify its content
            wb = _testDataProvider.WriteOutAndReadBack(wb);
            IName nm1 = wb.GetName("RangeTest1");
            Assert.IsTrue("RangeTest1".Equals(nm1.NameName), "Name is " + nm1.NameName);
            Assert.IsTrue((wb.GetSheetName(0) + "!$A$1:$L$41").Equals(nm1.RefersToFormula), "Reference is " + nm1.RefersToFormula);

            IName nm2 = wb.GetName("RangeTest2");
            Assert.IsTrue("RangeTest2".Equals(nm2.NameName), "Name is " + nm2.NameName);
            Assert.IsTrue((wb.GetSheetName(1) + "!$A$1:$O$21").Equals(nm2.RefersToFormula), "Reference is " + nm2.RefersToFormula);
        }

        /**
         * Test to see if the print areas can be retrieved/Created in memory
         */
        [Test]
        public void TestSinglePrintArea()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
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
        [Test]
        public void TestSinglePrintAreaWOSheet()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
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
        [Test]
        public void TestPrintAreaFile()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
            workbook.CreateSheet("Test Print Area");
            String sheetName = workbook.GetSheetName(0);


            String reference = "$A$1:$B$1";
            workbook.SetPrintArea(0, reference);

            workbook = _testDataProvider.WriteOutAndReadBack(workbook);

            String retrievedPrintArea = workbook.GetPrintArea(0);
            Assert.IsNotNull(retrievedPrintArea, "Print Area not defined for first sheet");
            Assert.AreEqual("'" + sheetName + "'!$A$1:$B$1", retrievedPrintArea, "References Match");
        }

        /**
         * Test to see if multiple print areas made it to the file
         */
        [Test]
        public void TestMultiplePrintAreaFile()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();

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

            // Check print areas After re-reading workbook
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
        [Test]
        public void TestPrintAreaCoords()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
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
        [Test]
        public void TestPrintAreaUnion()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
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
        [Test]
        public void TestPrintAreaRemove()
        {
            IWorkbook workbook = _testDataProvider.CreateWorkbook();
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
        [Test]
        public void TestMultipleNamedWrite()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();


            wb.CreateSheet("testSheet1");
            String sheetName = wb.GetSheetName(0);

            Assert.AreEqual("testSheet1", sheetName);

            //Creating new Named Range
            IName newNamedRange = wb.CreateName();

            newNamedRange.NameName = ("RangeTest");
            newNamedRange.RefersToFormula = (sheetName + "!$D$4:$E$8");

            //Creating another new Named Range
            IName newNamedRange2 = wb.CreateName();

            newNamedRange2.NameName = ("AnotherTest");
            newNamedRange2.RefersToFormula = (sheetName + "!$F$1:$G$6");

            wb.GetNameAt(0);

            wb = _testDataProvider.WriteOutAndReadBack(wb);
            IName nm = wb.GetName("RangeTest");
            Assert.IsTrue("RangeTest".Equals(nm.NameName), "Name is " + nm.NameName);
            Assert.IsTrue((wb.GetSheetName(0) + "!$D$4:$E$8").Equals(nm.RefersToFormula), "Reference is " + nm.RefersToFormula);

            nm = wb.GetName("AnotherTest");
            Assert.IsTrue("AnotherTest".Equals(nm.NameName), "Name is " + nm.NameName);
            Assert.IsTrue(newNamedRange2.RefersToFormula.Equals(nm.RefersToFormula), "Reference is " + nm.RefersToFormula);
        }
        /**
         * Verifies correct functioning for "single cell named range" (aka "named cell")
         */
        [Test]
        public void TestNamedCell_1()
        {

            // Setup for this Testcase
            String sheetName = "Test Named Cell";
            String cellName = "named_cell";
            String cellValue = "TEST Value";
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sheet = wb.CreateSheet(sheetName);
            ICreationHelper factory = wb.GetCreationHelper();
            sheet.CreateRow(0).CreateCell(0).SetCellValue(factory.CreateRichTextString(cellValue));

            // create named range for a single cell using areareference
            IName namedCell = wb.CreateName();
            namedCell.NameName = (cellName);
            String reference = "'" + sheetName + "'" + "!A1:A1";
            namedCell.RefersToFormula = (reference);

            // retrieve the newly Created named range
            IName aNamedCell = wb.GetName(cellName);
            Assert.IsNotNull(aNamedCell);

            // retrieve the cell at the named range and Test its contents
            RangeAddress aref = new RangeAddress(aNamedCell.RefersToFormula);
            Assert.IsTrue(aref.Height == 1 && aref.Width == 1, "Should be exactly 1 cell in the named cell :'" + cellName + "'");

            CellReference cref = new CellReference($"\'{aref.SheetName}\'!{aref.ToCell}");
            Assert.IsNotNull(cref);
            ISheet s = wb.GetSheet(cref.SheetName);
            Assert.IsNotNull(s);
            IRow r = sheet.GetRow(cref.Row);
            ICell c = r.GetCell(cref.Col);
            String contents = c.RichStringCellValue.String;
            Assert.AreEqual(contents, cellValue, "Contents of cell retrieved by its named reference");
        }

        /**
         * Verifies correct functioning for "single cell named range" (aka "named cell")
         */
        [Test]
        public void TestNamedCell_2()
        {

            // Setup for this Testcase
            String sname = "TestSheet", cname = "TestName", cvalue = "TestVal";
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICreationHelper factory = wb.GetCreationHelper();
            ISheet sheet = wb.CreateSheet(sname);
            sheet.CreateRow(0).CreateCell(0).SetCellValue(factory.CreateRichTextString(cvalue));

            // create named range for a single cell using cellreference
            IName namedCell = wb.CreateName();
            namedCell.NameName = (cname);
            String reference = sname + "!A1";
            namedCell.RefersToFormula = (reference);

            // retrieve the newly Created named range
            IName aNamedCell = wb.GetName(cname);
            Assert.IsNotNull(aNamedCell);

            // retrieve the cell at the named range and Test its contents
            CellReference cref = new CellReference(aNamedCell.RefersToFormula);
            Assert.IsNotNull(cref);
            ISheet s = wb.GetSheet(cref.SheetName);
            IRow r = sheet.GetRow(cref.Row);
            ICell c = r.GetCell(cref.Col);
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
         * This caused trouble for anything that requires {@link Name#getRefersToFormula()}
         * It is easy enough to re-create the the same data (by not Setting the formula). Excel
         * seems to gracefully remove this unInitialized name record.  It would be nice if POI
         * could do the same, but that would involve adjusting subsequent name indexes across
         * all formulas. <p/>
         *
         * For the moment, POI has been made to behave more sensibly with unInitialised name
         * records.
         */
        [Test]
        public void TestUnInitialisedNameGetRefersToFormula_bug46973()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IName n = wb.CreateName();
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
                    throw new AssertionException("Identified bug 46973");
                }
                throw e;
            }
            Assert.IsNull(formula);
            Assert.IsFalse(n.IsDeleted); // according to exact defInition of isDeleted()
        }
        [Test]
        public void TestDeletedCell()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IName n = wb.CreateName();
            n.NameName = ("MyName");
            // contrived example to expose bug:
            n.RefersToFormula = ("if(A1,\"#REF!\", \"\")");

            if (n.IsDeleted)
            {
                throw new AssertionException("Identified bug in recoginising formulas referring to deleted cells");
            }

        }
        [Test]
        public void TestFunctionNames()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IName n = wb.CreateName();
            Assert.IsFalse(n.IsFunctionName);

            n.SetFunction(false);
            Assert.IsFalse(n.IsFunctionName);

            n.SetFunction(true);
            Assert.IsTrue(n.IsFunctionName);

            n.SetFunction(false);
            Assert.IsFalse(n.IsFunctionName);
        }
        [Test]
        public void TestDefferedSetting()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            IName n1 = wb.CreateName();
            Assert.IsNull(n1.RefersToFormula);
            Assert.AreEqual("", n1.NameName);

            IName n2 = wb.CreateName();
            Assert.IsNull(n2.RefersToFormula);
            Assert.AreEqual("", n2.NameName);

            n1.NameName = ("sale_1");
            n1.RefersToFormula = ("10");

            n2.NameName = ("sale_2");
            n2.RefersToFormula = ("20");

            try
            {
                n2.NameName = ("sale_1");
                Assert.Fail("Expected exception");
            }
            catch (Exception e)
            {
                Assert.AreEqual("The workbook already contains this name: sale_1", e.Message);
            }

        }

        [Test]
        public void TestBug56930()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            // x1 on sheet1 defines "x=1"
            wb.CreateSheet("sheet1");
            IName x1 = wb.CreateName();
            x1.NameName = "x";
            x1.RefersToFormula = "1";
            x1.SheetIndex = wb.GetSheetIndex("sheet1");
            // x2 on sheet2 defines "x=2"
            wb.CreateSheet("sheet2");
            IName x2 = wb.CreateName();
            x2.NameName = "x";
            x2.RefersToFormula = "2";
            x2.SheetIndex = wb.GetSheetIndex("sheet2");
            IList<IName> names = wb.GetNames("x");
            Assert.AreEqual(2, names.Count, "Had: " + names);
            Assert.AreEqual("1", names[0].RefersToFormula);
            Assert.AreEqual("2", names[1].RefersToFormula);
            Assert.AreEqual("1", wb.GetName("x").RefersToFormula);
            wb.RemoveName("x");
            Assert.AreEqual("2", wb.GetName("x").RefersToFormula);

            wb.Close();
        }

        [Test]
        public void Test56781()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            IName name = wb.CreateName();
            foreach (String valid in Arrays.AsList(
                    "Hello",
                    "number1",
                    "_underscore"
                    //"p.e.r.o.i.d.s",
                    //"\\Backslash",
                    //"Backslash\\"
                    ))
            {
                name.NameName = valid;
            }

            try
            {
                name.NameName = "";
                Assert.Fail("expected exception: (blank)");
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Name cannot be blank", e.Message);
            }

            foreach (String invalid in Arrays.AsList(
                    "1number",
                    "Sheet1!A1",
                    "Exclamation!",
                    "Has Space",
                    "Colon:",
                    "A-Minus",
                    "A+Plus",
                    "Dollar$"))
            {
                try
                {
                    name.NameName = invalid;
                    Assert.Fail("expected exception: " + invalid);
                }
                catch (ArgumentException e)
                {
                    Assert.IsTrue(e.Message.StartsWith("Invalid name: '" + invalid + "'"), invalid);
                }
            }

        }

    }

}