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
    using System.Reflection;
    using NUnit.Framework;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using TestCases.SS.UserModel;

    /**
     * Tests various functionality having to do with {@link NPOI.SS.usermodel.Name}.
     *
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author ROMANL
     * @author Danny Mui (danny at muibros.com)
     * @author Amol S. Deshmukh &lt; amol at ap ache dot org &gt;
     */
    [TestFixture]
    public class TestHSSFName : BaseTestNamedRange
    {
        public TestHSSFName()
            : base(HSSFITestDataProvider.Instance)
        {

        }

        /**
         * For manipulating the internals of {@link HSSFName} during Testing.<br/>
         * Some Tests need a {@link NameRecord} with unusual state, not normally producible by POI.
         * This method achieves the aims at low cost without augmenting the POI usermodel api.
         * @return a reference to the wrapped {@link NameRecord} 
         */
        public static NameRecord GetNameRecord(IName definedName)
        {

            FieldInfo f;
            f = typeof(HSSFName).GetField("_definedNameRec", BindingFlags.Instance | BindingFlags.NonPublic);
            //f.SetAccessible(true);
            return (NameRecord)f.GetValue(definedName);
        }

        [Test]
        public void TestRepeatingRowsAndColumsNames()
        {
            // First Test that Setting RR&C for same sheet more than once only Creates a
            // single  Print_Titles built-in record
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)wb.CreateSheet("FirstSheet");

            // set repeating rows and columns twice for the first sheet
            CellRangeAddress cra = CellRangeAddress.ValueOf("A1:A3");
            for (int i = 0; i < 2; i++)
            {
                sheet.RepeatingColumns = (cra);
                sheet.RepeatingRows = (cra);
                sheet.CreateFreezePane(0, 3);
            }
            Assert.AreEqual(1, wb.NumberOfNames);
            IName nr1 = wb.GetNameAt(0);

            Assert.AreEqual("Print_Titles", nr1.NameName);
            // TODO - full column references not rendering properly, absolute markers not present either
            // assertEquals("FirstSheet!$A:$A,FirstSheet!$1:$3", nr1.getRefersToFormula());
            Assert.AreEqual("FirstSheet!A:A,FirstSheet!$A$1:$IV$3", nr1.RefersToFormula);

            // Save and re-open
            HSSFWorkbook nwb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            Assert.AreEqual(1, nwb.NumberOfNames);
            nr1 = nwb.GetNameAt(0);

            Assert.AreEqual("Print_Titles", nr1.NameName);
            Assert.AreEqual("FirstSheet!A:A,FirstSheet!$A$1:$IV$3", nr1.RefersToFormula);

            // check that Setting RR&C on a second sheet causes a new Print_Titles built-in
            // name to be Created
            sheet = (HSSFSheet)nwb.CreateSheet("SecondSheet");
            cra = CellRangeAddress.ValueOf("B1:C1");
            sheet.RepeatingColumns = (cra);
            sheet.RepeatingRows = (cra);

            Assert.AreEqual(2, nwb.NumberOfNames);
            IName nr2 = nwb.GetNameAt(1);

            Assert.AreEqual("Print_Titles", nr2.NameName);
            Assert.AreEqual("SecondSheet!B:C,SecondSheet!$A$1:$IV$1", nr2.RefersToFormula);
        }
        [Test]
        public void TestNamedRange()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("Simple.xls");

            //Creating new Named Range
            IName newNamedRange = wb.CreateName();

            //Getting Sheet Name for the reference
            String sheetName = wb.GetSheetName(0);

            //Setting its name
            newNamedRange.NameName = "RangeTest";
            //Setting its reference
            newNamedRange.RefersToFormula = sheetName + "!$D$4:$E$8";

            //Getting NAmed Range
            IName namedRange1 = wb.GetNameAt(0);
            //Getting it sheet name
            sheetName = namedRange1.SheetName;
            Assert.IsNotNull(sheetName);

            // sanity check
            SanityChecker c = new SanityChecker();
            c.CheckHSSFWorkbook(wb);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            IName nm = wb.GetNameAt(wb.GetNameIndex("RangeTest"));
            Assert.IsTrue("RangeTest".Equals(nm.NameName), "Name is " + nm.NameName);
            Assert.AreEqual(wb.GetSheetName(0) + "!$D$4:$E$8", nm.RefersToFormula);
        }

        /**
         * Reads an excel file already Containing a named range.
         * <p>
         * Addresses Bug <a href="http://issues.apache.org/bugzilla/Show_bug.cgi?id=9632" tarGet="_bug">#9632</a>
         */
        [Test]
        public void TestNamedRead()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("namedinput.xls");

            //Get index of the named range with the name = "NamedRangeName" , which was defined in input.xls as A1:D10
            int NamedRangeIndex = wb.GetNameIndex("NamedRangeName");

            //Getting NAmed Range
            IName namedRange1 = wb.GetNameAt(NamedRangeIndex);
            String sheetName = wb.GetSheetName(0);

            //Getting its reference
            String reference = namedRange1.RefersToFormula;

            Assert.AreEqual(sheetName + "!$A$1:$D$10", reference);

            IName namedRange2 = wb.GetNameAt(1);

            Assert.AreEqual(sheetName + "!$D$17:$G$27", namedRange2.RefersToFormula);
            Assert.AreEqual("SecondNamedRange", namedRange2.NameName);
        }

        /**
         * Reads an excel file already Containing a named range and updates it
         * <p>
         * Addresses Bug <a href="http://issues.apache.org/bugzilla/Show_bug.cgi?id=16411" tarGet="_bug">#16411</a>
         */
        [Test]
        public void TestNamedReadModify()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("namedinput.xls");

            IName name = wb.GetNameAt(0);
            String sheetName = wb.GetSheetName(0);

            Assert.AreEqual(sheetName + "!$A$1:$D$10", name.RefersToFormula);

            name = wb.GetNameAt(1);
            String newReference = sheetName + "!$A$1:$C$36";

            name.RefersToFormula = newReference;
            Assert.AreEqual(newReference, name.RefersToFormula);
        }

        /**
         * Test to see if the print area can be retrieved from an excel Created file
         */
        [Test]
        public void TestPrintAreaFileRead()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithPrintArea.xls");

            String sheetName = workbook.GetSheetName(0);
            String reference = sheetName + "!$A$1:$C$5";

            Assert.AreEqual(reference, workbook.GetPrintArea(0));
        }

        [Test]
        [Obsolete]
        public void TestDeletedReference()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("24207.xls");
            Assert.AreEqual(2, wb.NumberOfNames);

            IName name1 = wb.GetNameAt(0);
            Assert.AreEqual("a", name1.NameName);
            Assert.AreEqual("Sheet1!$A$1", name1.RefersToFormula);
            new AreaReference(name1.RefersToFormula);
            Assert.IsTrue(true, "Successfully constructed first reference");

            IName name2 = wb.GetNameAt(1);
            Assert.AreEqual("b", name2.NameName);
            Assert.AreEqual("Sheet1!#REF!", name2.RefersToFormula);
            Assert.IsTrue(name2.IsDeleted);
            try
            {
                new AreaReference(name2.RefersToFormula);
                Assert.Fail("attempt to supply an invalid reference to AreaReference constructor results in exception");
            }
            catch (ArgumentException)
            { // TODO - use a different exception for this condition
                // expected during successful Test
            }
        }

        /**
         * When Setting A1 type of references with HSSFName.SetRefersToFormula
         * must set the type of operands to Ptg.CLASS_REF,
         * otherwise Created named don't appear in the drop-down to the left of formula bar in Excel
         */
        [Test]
        public void TestTypeOfRootPtg()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("CSCO");

            Ptg[] ptgs = HSSFFormulaParser.Parse("CSCO!$E$71", wb, FormulaType.NamedRange, 0);
            foreach (Ptg ptg in ptgs)
            {
                Assert.AreEqual('R', ptg.RVAType);
            }
        }
    }
}
