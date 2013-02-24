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
namespace NPOI.XSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS;
    using NPOI.XSSF;
    using NUnit.Framework;
    using NPOI.SS.Util;
    using NPOI.OpenXmlFormats.Spreadsheet;
    [TestFixture]
    public class TestXSSFBugs : BaseTestBugzillaIssues
    {

        public TestXSSFBugs()
            : base(XSSFITestDataProvider.instance)
        {

        }

        /**
         * Test writing a file with large number of unique strings,
         * open resulting file in Excel to check results!
         */
        [Test]
        public void Test15375_2()
        {
            BaseTest15375(1000);
        }

        /**
         * Named ranges had the right reference, but
         *  the wrong sheet name
         */
        [Test]
        public void Test45430()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("45430.xlsx");
            Assert.IsFalse(wb.IsMacroEnabled());
            Assert.AreEqual(3, wb.NumberOfNames);

            XSSFName name = (wb.GetNameAt(0) as XSSFName);
            Assert.AreEqual(0, name.GetCTName().localSheetId);
            Assert.IsFalse(name.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual("SheetA!$A$1", name.RefersToFormula);
            Assert.AreEqual("SheetA", name.SheetName);

            name = (wb.GetNameAt(1) as XSSFName);
            Assert.AreEqual(0, name.GetCTName().localSheetId);
            Assert.IsFalse(name.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual("SheetB!$A$1", wb.GetNameAt(1).RefersToFormula);
            Assert.AreEqual("SheetB", wb.GetNameAt(1).SheetName);

            name = (wb.GetNameAt(2) as XSSFName);
            Assert.AreEqual(0, name.GetCTName().localSheetId);
            Assert.IsFalse(name.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual("SheetC!$A$1", wb.GetNameAt(2).RefersToFormula);
            Assert.AreEqual("SheetC", wb.GetNameAt(2).SheetName);

            // Save and re-load, still there
            XSSFWorkbook nwb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            Assert.AreEqual(3, nwb.NumberOfNames);
            Assert.AreEqual("SheetA!$A$1", nwb.GetNameAt(0).RefersToFormula);
        }

        /**
         * We should carry vba macros over After save
         */
        [Test]
        public void Test45431()
        {
            throw new NotImplementedException();
            //        XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("45431.xlsm");
            //        OPCPackage pkg = wb.GetPackage();
            //        Assert.IsTrue(wb.IsMacroEnabled());

            //        // Check the various macro related bits can be found
            //        PackagePart vba = pkg.GetPart(
            //                PackagingURIHelper.CreatePartName("/xl/vbaProject.bin")
            //        );
            //        Assert.IsNotNull(vba);
            //        // And the Drawing bit
            //        PackagePart drw = pkg.GetPart(
            //                PackagingURIHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            //        );
            //        Assert.IsNotNull(drw);


            //        // Save and re-open, both still there
            //        XSSFWorkbook nwb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //        OPCPackage nPkg = nwb.GetPackage();
            //        Assert.IsTrue(nwb.IsMacroEnabled());

            //        vba = nPkg.GetPart(
            //                PackagingURIHelper.CreatePartName("/xl/vbaProject.bin")
            //        );
            //        Assert.IsNotNull(vba);
            //        drw = nPkg.GetPart(
            //                PackagingURIHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            //        );
            //        Assert.IsNotNull(drw);

            //        // And again, just to be sure
            //        nwb = XSSFTestDataSamples.WriteOutAndReadBack(nwb);
            //        nPkg = nwb.GetPackage();
            //        Assert.IsTrue(nwb.IsMacroEnabled());

            //        vba = nPkg.GetPart(
            //                PackagingURIHelper.CreatePartName("/xl/vbaProject.bin")
            //        );
            //        Assert.IsNotNull(vba);
            //        drw = nPkg.GetPart(
            //                PackagingURIHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            //        );
            //        Assert.IsNotNull(drw);
        }
        [Test]
        public void Test47504()
        {
            throw new NotImplementedException();
            //        XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("47504.xlsx");
            //        Assert.AreEqual(1, wb.GetNumberOfSheets());
            //        XSSFSheet sh = wb.GetSheetAt(0);
            //        XSSFDrawing Drawing = sh.CreateDrawingPatriarch();
            //        List<POIXMLDocumentPart> rels = Drawing.GetRelations();
            //        Assert.AreEqual(1, rels.Count);
            //        Assert.AreEqual("Sheet1!A1", rels.Get(0).GetPackageRelationship().GetTargetURI().GetFragment());

            //        // And again, just to be sure
            //        wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //        Assert.AreEqual(1, wb.GetNumberOfSheets());
            //        sh = wb.GetSheetAt(0);
            //        Drawing = sh.CreateDrawingPatriarch();
            //        rels = Drawing.GetRelations();
            //        Assert.AreEqual(1, rels.Count);
            //        Assert.AreEqual("Sheet1!A1", rels.Get(0).GetPackageRelationship().GetTargetURI().GetFragment());
        }

        /**
         * Excel will sometimes write a button with a textbox
         *  Containing &gt;br&lt; (not closed!).
         * Clearly Excel shouldn't do this, but Test that we can
         *  read the file despite the naughtyness
         */
        [Test]
        public void Test49020()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("BrNotClosed.xlsx");
        }

        /**
         * ensure that CTPhoneticPr is loaded by the ooxml Test suite so that it is included in poi-ooxml-schemas
         */
        [Test]
        public void Test49325()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("49325.xlsx");
            CT_Worksheet sh = (wb.GetSheetAt(0) as XSSFSheet).GetCTWorksheet();
            Assert.IsNotNull(sh.phoneticPr);
        }

        /**
         * Names which are defined with a Sheet
         *  should return that sheet index properly 
         */
        [Test]
        public void Test48923()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48923.xlsx");
            //       Assert.AreEqual(4, wb.GetNumberOfNames());

            //       Name b1 = wb.GetName("NameB1");
            //       Name b2 = wb.GetName("NameB2");
            //       Name sheet2 = wb.GetName("NameSheet2");
            //       Name Test = wb.GetName("Test");

            //       Assert.IsNotNull(b1);
            //       Assert.AreEqual("NameB1", b1.GetNameName());
            //       Assert.AreEqual("Sheet1", b1.GetSheetName());
            //       Assert.AreEqual(-1, b1.GetSheetIndex());

            //       Assert.IsNotNull(b2);
            //       Assert.AreEqual("NameB2", b2.GetNameName());
            //       Assert.AreEqual("Sheet1", b2.GetSheetName());
            //       Assert.AreEqual(-1, b2.GetSheetIndex());

            //       Assert.IsNotNull(sheet2);
            //       Assert.AreEqual("NameSheet2", sheet2.GetNameName());
            //       Assert.AreEqual("Sheet2", sheet2.GetSheetName());
            //       Assert.AreEqual(-1, sheet2.GetSheetIndex());

            //       Assert.IsNotNull(test);
            //       Assert.AreEqual("Test", Test.GetNameName());
            //       Assert.AreEqual("Sheet1", Test.GetSheetName());
            //       Assert.AreEqual(-1, Test.GetSheetIndex());
        }

        /**
         * Problem with Evaluation formulas due to
         *  NameXPtgs.
         * Blows up on:
         *   IF(B6= (ROUNDUP(B6,0) + ROUNDDOWN(B6,0))/2, MROUND(B6,2),ROUND(B6,0))
         * 
         * TODO: delete this Test case when MROUND and VAR are implemented
         */
        [Test]
        public void Test48539()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48539.xlsx");
            //       Assert.AreEqual(3, wb.GetNumberOfSheets());

            //       // Try each cell individually
            //       XSSFFormulaEvaluator eval = new XSSFFormulaEvaluator(wb);
            //       for(int i=0; i<wb.GetNumberOfSheets(); i++) {
            //          Sheet s = wb.GetSheetAt(i);
            //          foreach(Row r in s) {
            //             foreach(Cell c in r) {
            //                if(c.GetCellType() == Cell.CELL_TYPE_FORMULA) {
            //                    CellValue cv = Eval.evaluate(c);
            //                    if(cv.GetCellType() == Cell.CELL_TYPE_NUMERIC) {
            //                        // assert that the calculated value agrees with
            //                        // the cached formula result calculated by Excel
            //                        double cachedFormulaResult = c.GetNumericCellValue();
            //                        double EvaluatedFormulaResult = cv.GetNumberValue();
            //                        Assert.AreEqual(c.GetCellFormula(), cachedFormulaResult, EvaluatedFormulaResult, 1E-7);
            //                    }
            //                }
            //             }
            //          }
            //       }

            //       // Now all of them
            //        XSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
        }

        /**
         * Foreground colours should be found even if
         *  a theme is used 
         */
        [Test]
        public void Test48779()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48779.xlsx");
            //       XSSFCell cell = wb.GetSheetAt(0).GetRow(0).GetCell(0);
            //       XSSFCellStyle cs = cell.GetCellStyle();

            //       Assert.IsNotNull(cs);
            //       Assert.AreEqual(1, cs.GetIndex());

            //       // Look at the low level xml elements
            //       Assert.AreEqual(2, cs.GetCoreXf().GetFillId());
            //       Assert.AreEqual(0, cs.GetCoreXf().GetXfId());
            //       Assert.AreEqual(true, cs.GetCoreXf().GetApplyFill());

            //       XSSFCellFill fg = wb.GetStylesSource().GetFillAt(2);
            //       Assert.AreEqual(0, fg.GetFillForegroundColor().GetIndexed());
            //       Assert.AreEqual(0.0, fg.GetFillForegroundColor().GetTint());
            //       Assert.AreEqual("FFFF0000", fg.GetFillForegroundColor().GetARGBHex());
            //       Assert.AreEqual(64, fg.GetFillBackgroundColor().GetIndexed());

            //       // Now look higher up
            //       Assert.IsNotNull(cs.GetFillForegroundXSSFColor());
            //       Assert.AreEqual(0, cs.GetFillForegroundColor());
            //       Assert.AreEqual("FFFF0000", cs.GetFillForegroundXSSFColor().GetARGBHex());
            //       Assert.AreEqual("FFFF0000", cs.GetFillForegroundColorColor().GetARGBHex());

            //       Assert.IsNotNull(cs.GetFillBackgroundColor());
            //       Assert.AreEqual(64, cs.GetFillBackgroundColor());
            //       Assert.AreEqual(null, cs.GetFillBackgroundXSSFColor().GetARGBHex());
            //       Assert.AreEqual(null, cs.GetFillBackgroundColorColor().GetARGBHex());
        }

        /**
         * With HSSF, if you create a font, don't change it, and
         *  create a 2nd, you really do Get two fonts that you 
         *  can alter as and when you want.
         * With XSSF, that wasn't the case, but this verfies
         *  that it now is again
         */
        [Test]
        public void Test48718()
        {
            throw new NotImplementedException();
            //       // Verify the HSSF behaviour
            //       // Then ensure the same for XSSF
            //       Workbook[] wbs = new Workbook[] {
            //             new HSSFWorkbook(),
            //             new XSSFWorkbook()
            //       };
            //       int[] InitialFonts = new int[] { 4, 1 };
            //       for(int i=0; i<wbs.Length; i++) {
            //          Workbook wb = wbs[i];
            //          int startingFonts = InitialFonts[i];

            //          Assert.AreEqual(startingFonts, wb.GetNumberOfFonts());

            //          // Get a font, and slightly change it
            //          Font a = wb.CreateFont();
            //          Assert.AreEqual(startingFonts+1, wb.GetNumberOfFonts());
            //          a.SetFontHeightInPoints((short)23);
            //          Assert.AreEqual(startingFonts+1, wb.GetNumberOfFonts());

            //          // Get two more, unChanged
            //          Font b = wb.CreateFont();
            //          Assert.AreEqual(startingFonts+2, wb.GetNumberOfFonts());
            //          Font c = wb.CreateFont();
            //          Assert.AreEqual(startingFonts+3, wb.GetNumberOfFonts());
            //       }
        }

        /**
         * Ensure General and @ format are working properly
         *  for integers 
         */
        [Test]
        public void Test47490()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("GeneralFormatTests.xlsx");
            //       Sheet s = wb.GetSheetAt(1);
            //       Row r;
            //       DataFormatter df = new DataFormatter();

            //       r = s.GetRow(1);
            //       Assert.AreEqual(1.0, r.GetCell(2).GetNumericCellValue());
            //       Assert.AreEqual("General", r.GetCell(2).GetCellStyle().GetDataFormatString());
            //       Assert.AreEqual("1", df.formatCellValue(r.GetCell(2)));
            //       Assert.AreEqual("1", df.formatRawCellContents(1.0, -1, "@"));
            //       Assert.AreEqual("1", df.formatRawCellContents(1.0, -1, "General"));

            //       r = s.GetRow(2);
            //       Assert.AreEqual(12.0, r.GetCell(2).GetNumericCellValue());
            //       Assert.AreEqual("General", r.GetCell(2).GetCellStyle().GetDataFormatString());
            //       Assert.AreEqual("12", df.formatCellValue(r.GetCell(2)));
            //       Assert.AreEqual("12", df.formatRawCellContents(12.0, -1, "@"));
            //       Assert.AreEqual("12", df.formatRawCellContents(12.0, -1, "General"));

            //       r = s.GetRow(3);
            //       Assert.AreEqual(123.0, r.GetCell(2).GetNumericCellValue());
            //       Assert.AreEqual("General", r.GetCell(2).GetCellStyle().GetDataFormatString());
            //       Assert.AreEqual("123", df.formatCellValue(r.GetCell(2)));
            //       Assert.AreEqual("123", df.formatRawCellContents(123.0, -1, "@"));
            //       Assert.AreEqual("123", df.formatRawCellContents(123.0, -1, "General"));
        }

        /**
         * Ensures that XSSF and HSSF agree with each other,
         *  and with the docs on when fetching the wrong
         *  kind of value from a Formula cell
         */
        [Test]
        public void Test47815()
        {
            throw new NotImplementedException();
            //       Workbook[] wbs = new Workbook[] {
            //             new HSSFWorkbook(),
            //             new XSSFWorkbook()
            //       };
            //       foreach(Workbook wb in wbs) {
            //          Sheet s = wb.CreateSheet();
            //          Row r = s.CreateRow(0);

            //          // Setup
            //          Cell cn = r.CreateCell(0, Cell.CELL_TYPE_NUMERIC);
            //          cn.SetCellValue(1.2);
            //          Cell cs = r.CreateCell(1, Cell.CELL_TYPE_STRING);
            //          cs.SetCellValue("Testing");

            //          Cell cfn = r.CreateCell(2, Cell.CELL_TYPE_FORMULA);
            //          cfn.SetCellFormula("A1");  
            //          Cell cfs = r.CreateCell(3, Cell.CELL_TYPE_FORMULA);
            //          cfs.SetCellFormula("B1");

            //          FormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
            //          Assert.AreEqual(Cell.CELL_TYPE_NUMERIC, fe.Evaluate(cfn).GetCellType());
            //          Assert.AreEqual(Cell.CELL_TYPE_STRING, fe.Evaluate(cfs).GetCellType());
            //          fe.EvaluateFormulaCell(cfn);
            //          fe.EvaluateFormulaCell(cfs);

            //          // Now Test
            //          Assert.AreEqual(Cell.CELL_TYPE_NUMERIC, cn.GetCellType());
            //          Assert.AreEqual(Cell.CELL_TYPE_STRING, cs.GetCellType());
            //          Assert.AreEqual(Cell.CELL_TYPE_FORMULA, cfn.GetCellType());
            //          Assert.AreEqual(Cell.CELL_TYPE_NUMERIC, cfn.GetCachedFormulaResultType());
            //          Assert.AreEqual(Cell.CELL_TYPE_FORMULA, cfs.GetCellType());
            //          Assert.AreEqual(Cell.CELL_TYPE_STRING, cfs.GetCachedFormulaResultType());

            //          // Different ways of retrieving
            //          Assert.AreEqual(1.2, cn.GetNumericCellValue());
            //          try {
            //             cn.GetRichStringCellValue();
            //             Assert.Fail();
            //          } catch(InvalidOperationException e) {}

            //          Assert.AreEqual("Testing", cs.GetStringCellValue());
            //          try {
            //             cs.GetNumericCellValue();
            //             Assert.Fail();
            //          } catch(InvalidOperationException e) {}

            //          Assert.AreEqual(1.2, cfn.GetNumericCellValue());
            //          try {
            //             cfn.GetRichStringCellValue();
            //             Assert.Fail();
            //          } catch(InvalidOperationException e) {}

            //          Assert.AreEqual("Testing", cfs.GetStringCellValue());
            //          try {
            //             cfs.GetNumericCellValue();
            //             Assert.Fail();
            //          } catch(InvalidOperationException e) {}
            //       }
        }

        /**
         * A problem file from a non-standard source (a scientific instrument that saves its
         * output as an .xlsx file) that have two issues:
         * 1. The Content Type part name is lower-case:  [content_types].xml
         * 2. The file appears to use backslashes as path separators
         *
         * The OPC spec tolerates both of these peculiarities, so does POI
         */
        [Test]
        public void Test49609()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("49609.xlsx");
            Assert.AreEqual("FAM", wb.GetSheetName(0));
            Assert.AreEqual("Cycle", wb.GetSheetAt(0).GetRow(0).GetCell(1).StringCellValue);

        }
        [Test]
        public void Test49783()
        {
            throw new NotImplementedException();
            //        Workbook wb =  XSSFTestDataSamples.OpenSampleWorkbook("49783.xlsx");
            //        Sheet sheet = wb.GetSheetAt(0);
            //        FormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            //        Cell cell;

            //        cell = sheet.GetRow(0).GetCell(0);
            //        Assert.AreEqual("#REF!*#REF!", cell.GetCellFormula());
            //        Assert.AreEqual(Cell.CELL_TYPE_ERROR, Evaluator.evaluateInCell(cell).GetCellType());
            //        Assert.AreEqual("#REF!", FormulaError.forInt(cell.GetErrorCellValue()).GetString());

            //        Name nm1 = wb.GetName("sale_1");
            //        Assert.IsNotNull("name sale_1 should be present", nm1);
            //        Assert.AreEqual("Sheet1!#REF!", nm1.GetRefersToFormula());
            //        Name nm2 = wb.GetName("sale_2");
            //        Assert.IsNotNull("name sale_2 should be present", nm2);
            //        Assert.AreEqual("Sheet1!#REF!", nm2.GetRefersToFormula());

            //        cell = sheet.GetRow(1).GetCell(0);
            //        Assert.AreEqual("sale_1*sale_2", cell.GetCellFormula());
            //        Assert.AreEqual(Cell.CELL_TYPE_ERROR, Evaluator.evaluateInCell(cell).GetCellType());
            //        Assert.AreEqual("#REF!", FormulaError.forInt(cell.GetErrorCellValue()).GetString());
        }

        /**
         * Creating a rich string of "hello world" and Applying
         *  a font to characters 1-5 means we have two strings,
         *  "hello" and " world". As such, we need to apply
         *  preserve spaces to the 2nd bit, lest we end up
         *  with something like "helloworld" !
         */
        [Test]
        public void Test49941()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = new XSSFWorkbook();
            //       XSSFSheet s = wb.CreateSheet();
            //       XSSFRow r = s.CreateRow(0);
            //       XSSFCell c = r.CreateCell(0);

            //       // First without fonts
            //       c.SetCellValue(
            //             new XSSFRichTextString(" with spaces ")
            //       );
            //       Assert.AreEqual(" with spaces ", c.GetRichStringCellValue().ToString());
            //       Assert.AreEqual(0, c.GetRichStringCellValue().GetCTRst().sizeOfRArray());
            //       Assert.AreEqual(true, c.GetRichStringCellValue().GetCTRst().IsSetT());
            //       // Should have the preserve Set
            //       Assert.AreEqual(
            //             1,
            //             c.GetRichStringCellValue().GetCTRst().xgetT().GetDomNode().GetAttributes().GetLength()
            //       );
            //       Assert.AreEqual(
            //             "preserve",
            //             c.GetRichStringCellValue().GetCTRst().xgetT().GetDomNode().GetAttributes().item(0).GetNodeValue()
            //       );

            //       // Save and check
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       s = wb.GetSheetAt(0);
            //       r = s.GetRow(0);
            //       c = r.GetCell(0);
            //       Assert.AreEqual(" with spaces ", c.GetRichStringCellValue().ToString());
            //       Assert.AreEqual(0, c.GetRichStringCellValue().GetCTRst().sizeOfRArray());
            //       Assert.AreEqual(true, c.GetRichStringCellValue().GetCTRst().IsSetT());

            //       // Change the string
            //       c.SetCellValue(
            //             new XSSFRichTextString("hello world")
            //       );
            //       Assert.AreEqual("hello world", c.GetRichStringCellValue().ToString());
            //       // Won't have preserve
            //       Assert.AreEqual(
            //             0,
            //             c.GetRichStringCellValue().GetCTRst().xgetT().GetDomNode().GetAttributes().GetLength()
            //       );

            //       // Apply a font
            //       XSSFFont f = wb.CreateFont();
            //       f.SetBold(true);
            //       c.GetRichStringCellValue().ApplyFont(0, 5, f);
            //       Assert.AreEqual("hello world", c.GetRichStringCellValue().ToString());
            //       // Does need preserving on the 2nd part
            //       Assert.AreEqual(2, c.GetRichStringCellValue().GetCTRst().sizeOfRArray());
            //       Assert.AreEqual(
            //             0,
            //             c.GetRichStringCellValue().GetCTRst().GetRArray(0).xgetT().GetDomNode().GetAttributes().GetLength()
            //       );
            //       Assert.AreEqual(
            //             1,
            //             c.GetRichStringCellValue().GetCTRst().GetRArray(1).xgetT().GetDomNode().GetAttributes().GetLength()
            //       );
            //       Assert.AreEqual(
            //             "preserve",
            //             c.GetRichStringCellValue().GetCTRst().GetRArray(1).xgetT().GetDomNode().GetAttributes().item(0).GetNodeValue()
            //       );

            //       // Save and check
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       s = wb.GetSheetAt(0);
            //       r = s.GetRow(0);
            //       c = r.GetCell(0);
            //       Assert.AreEqual("hello world", c.GetRichStringCellValue().ToString());
        }

        /**
         * Repeatedly writing the same file which has styles
         * TODO Currently failing
         */
        public void DISABLEDtest49940()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("styles.xlsx");
            //       Assert.AreEqual(3, wb.GetNumberOfSheets());
            //       Assert.AreEqual(10, wb.GetStylesSource().GetNumCellStyles());

            //       MemoryStream b1 = new MemoryStream();
            //       MemoryStream b2 = new MemoryStream();
            //       MemoryStream b3 = new MemoryStream();
            //       wb.Write(b1);
            //       wb.Write(b2);
            //       wb.Write(b3);

            //       foreach(byte[] data in new byte[][] {
            //             b1.ToArray(), b2.ToArray(), b3.ToArray()
            //       }) {
            //          MemoryStream bais = new MemoryStream(data);
            //          wb = new XSSFWorkbook(bais);
            //          Assert.AreEqual(3, wb.GetNumberOfSheets());
            //          Assert.AreEqual(10, wb.GetStylesSource().GetNumCellStyles());
            //       }
        }

        /**
         * Various ways of removing a cell formula should all zap
         *  the calcChain entry.
         */
        [Test]
        public void Test49966()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("shared_formulas.xlsx");
            //       XSSFSheet sheet = wb.GetSheetAt(0);

            //       // CalcChain has lots of entries
            //       CalculationChain cc = wb.GetCalculationChain();
            //       Assert.AreEqual("A2", cc.GetCTCalcChain().GetCArray(0).GetR());
            //       Assert.AreEqual("A3", cc.GetCTCalcChain().GetCArray(1).GetR());
            //       Assert.AreEqual("A4", cc.GetCTCalcChain().GetCArray(2).GetR());
            //       Assert.AreEqual("A5", cc.GetCTCalcChain().GetCArray(3).GetR());
            //       Assert.AreEqual("A6", cc.GetCTCalcChain().GetCArray(4).GetR());
            //       Assert.AreEqual("A7", cc.GetCTCalcChain().GetCArray(5).GetR());
            //       Assert.AreEqual("A8", cc.GetCTCalcChain().GetCArray(6).GetR());
            //       Assert.AreEqual(40, cc.GetCTCalcChain().sizeOfCArray());

            //       // Try various ways of changing the formulas
            //       // If it stays a formula, chain entry should remain
            //       // Otherwise should go
            //       sheet.GetRow(1).GetCell(0).SetCellFormula("A1"); // stay
            //       sheet.GetRow(2).GetCell(0).SetCellFormula(null);  // go
            //       sheet.GetRow(3).GetCell(0).SetCellType(Cell.CELL_TYPE_FORMULA); // stay
            //       sheet.GetRow(4).GetCell(0).SetCellType(Cell.CELL_TYPE_STRING);  // go
            //       sheet.GetRow(5).RemoveCell(
            //             sheet.GetRow(5).GetCell(0)  // go
            //       );
            //        sheet.GetRow(6).GetCell(0).SetCellType(Cell.CELL_TYPE_BLANK);  // go
            //        sheet.GetRow(7).GetCell(0).SetCellValue((String)null);  // go

            //       // Save and check
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       Assert.AreEqual(35, cc.GetCTCalcChain().sizeOfCArray());

            //       cc = wb.GetCalculationChain();
            //       Assert.AreEqual("A2", cc.GetCTCalcChain().GetCArray(0).GetR());
            //       Assert.AreEqual("A4", cc.GetCTCalcChain().GetCArray(1).GetR());
            //       Assert.AreEqual("A9", cc.GetCTCalcChain().GetCArray(2).GetR());

        }
        [Test]
        public void Test49156()
        {
            throw new NotImplementedException();
            //        Workbook wb = XSSFTestDataSamples.OpenSampleWorkbook("49156.xlsx");
            //        FormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            //        Sheet sheet = wb.GetSheetAt(0);
            //        foreach(Row row in sheet){
            //            foreach(Cell cell in row){
            //                if(cell.GetCellType() == Cell.CELL_TYPE_FORMULA){
            //                    formulaEvaluator.EvaluateInCell(cell); // caused NPE on some cells
            //                }
            //            }
            //        }
        }

        /**
         * Newlines are valid characters in a formula
         */
        [Test]
        public void Test50440And51875()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("NewlineInFormulas.xlsx");
            ISheet s = wb.GetSheetAt(0);
            ICell c = s.GetRow(0).GetCell(0);

            Assert.AreEqual("SUM(\n1,2\n)", c.CellFormula);
            Assert.AreEqual(3.0, c.NumericCellValue);

            IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            formulaEvaluator.EvaluateFormulaCell(c);

            Assert.AreEqual("SUM(\n1,2\n)", c.CellFormula);
            Assert.AreEqual(3.0, c.NumericCellValue);

            // For 51875
            ICell b3 = s.GetRow(2).GetCell(1);
            formulaEvaluator.EvaluateFormulaCell(b3);
            Assert.AreEqual("B1+B2", b3.CellFormula); // The newline is lost for shared formulas
            Assert.AreEqual(3.0, b3.NumericCellValue);
        }

        /**
         * Moving a cell comment from one cell to another
         */
        [Test]
        public void Test50795()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50795.xlsx");
            //       XSSFSheet sheet = wb.GetSheetAt(0);
            //       XSSFRow row = sheet.GetRow(0);

            //       XSSFCell cellWith = row.GetCell(0);
            //       XSSFCell cellWithoutComment = row.GetCell(1);

            //       Assert.IsNotNull(cellWith.GetCellComment());
            //       Assert.IsNull(cellWithoutComment.GetCellComment());

            //       String exp = "\u0410\u0432\u0442\u043e\u0440:\ncomment";
            //       XSSFComment comment = cellWith.GetCellComment();
            //       Assert.AreEqual(exp, comment.GetString().GetString());


            //       // Check we can write it out and read it back as-is
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       sheet = wb.GetSheetAt(0);
            //       row = sheet.GetRow(0);
            //       cellWith = row.GetCell(0);
            //       cellWithoutComment = row.GetCell(1);

            //       // Double check things are as expected
            //       Assert.IsNotNull(cellWith.GetCellComment());
            //       Assert.IsNull(cellWithoutComment.GetCellComment());
            //       comment = cellWith.GetCellComment();
            //       Assert.AreEqual(exp, comment.GetString().GetString());


            //       // Move the comment
            //       cellWithoutComment.SetCellComment(comment);


            //       // Write out and re-check
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       sheet = wb.GetSheetAt(0);
            //       row = sheet.GetRow(0);

            //       // Ensure it swapped over
            //       cellWith = row.GetCell(0);
            //       cellWithoutComment = row.GetCell(1);
            //       Assert.IsNull(cellWith.GetCellComment());
            //       Assert.IsNotNull(cellWithoutComment.GetCellComment());

            //       comment = cellWithoutComment.GetCellComment();
            //       Assert.AreEqual(exp, comment.GetString().GetString());
        }

        /**
         * When the cell background colour is Set with one of the first
         *  two columns of the theme colour palette, the colours are 
         *  shades of white or black.
         * For those cases, ensure we don't break on Reading the colour
         */
        [Test]
        public void Test50299()
        {
            throw new NotImplementedException();
            //       Workbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50299.xlsx");

            //       // Check all the colours
            //       for(int sn=0; sn<wb.GetNumberOfSheets(); sn++) {
            //          Sheet s = wb.GetSheetAt(sn);
            //          foreach(Row r in s) {
            //             foreach(Cell c in r) {
            //                CellStyle cs = c.GetCellStyle();
            //                if(cs != null) {
            //                   cs.GetFillForegroundColor();
            //                }
            //             }
            //          }
            //       }

            //       // Check one bit in detail
            //       // Check that we Get back foreground=0 for the theme colours,
            //       //  and background=64 for the auto colouring
            //       Sheet s = wb.GetSheetAt(0);
            //       Assert.AreEqual(0,  s.GetRow(0).GetCell(8).GetCellStyle().GetFillForegroundColor());
            //       Assert.AreEqual(64, s.GetRow(0).GetCell(8).GetCellStyle().GetFillBackgroundColor());
            //       Assert.AreEqual(0,  s.GetRow(1).GetCell(8).GetCellStyle().GetFillForegroundColor());
            //       Assert.AreEqual(64, s.GetRow(1).GetCell(8).GetCellStyle().GetFillBackgroundColor());
        }

        /**
         * Excel .xls style indexed colours in a .xlsx file
         */
        [Test]
        public void Test50786()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50786-indexed_colours.xlsx");
            //       XSSFSheet s = wb.GetSheetAt(0);
            //       XSSFRow r = s.GetRow(2);

            //       // Check we have the right cell
            //       XSSFCell c = r.GetCell(1);
            //       Assert.AreEqual("test\u00a0", c.GetRichStringCellValue().GetString());

            //       // It should be light green
            //       XSSFCellStyle cs = c.GetCellStyle();
            //       Assert.AreEqual(42, cs.GetFillForegroundColor());
            //       Assert.AreEqual(42, cs.GetFillForegroundColorColor().GetIndexed());
            //       Assert.IsNotNull(cs.GetFillForegroundColorColor().GetRgb());
            //       Assert.AreEqual("FFCCFFCC", cs.GetFillForegroundColorColor().GetARGBHex());
        }

        /**
         * If the border colours are Set with themes, then we 
         *  should still be able to Get colours
         */
        [Test]
        public void Test50846()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50846-border_colours.xlsx");

            //       XSSFSheet sheet = wb.GetSheetAt(0);
            //       XSSFRow row = sheet.GetRow(0);

            //       // Border from a theme, brown
            //       XSSFCell cellT = row.GetCell(0);
            //       XSSFCellStyle styleT = cellT.GetCellStyle();
            //       XSSFColor colorT = styleT.GetBottomBorderXSSFColor();

            //       Assert.AreEqual(5, colorT.GetTheme());
            //       Assert.AreEqual("FFC0504D", colorT.GetARGBHex());

            //       // Border from a style direct, red
            //       XSSFCell cellS = row.GetCell(1);
            //       XSSFCellStyle styleS = cellS.GetCellStyle();
            //       XSSFColor colorS = styleS.GetBottomBorderXSSFColor();

            //       Assert.AreEqual(0, colorS.GetTheme());
            //       Assert.AreEqual("FFFF0000", colorS.GetARGBHex());
        }

        /**
         * Fonts where their colours come from the theme rather
         *  then being Set explicitly still should allow the
         *  fetching of the RGB.
         */
        [Test]
        public void Test50784()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50784-font_theme_colours.xlsx");
            //       XSSFSheet s = wb.GetSheetAt(0);
            //       XSSFRow r = s.GetRow(0);

            //       // Column 1 has a font with regular colours
            //       XSSFCell cr = r.GetCell(1);
            //       XSSFFont fr = wb.GetFontAt( cr.GetCellStyle().GetFontIndex() );
            //       XSSFColor colr =  fr.GetXSSFColor();
            //       // No theme, has colours
            //       Assert.AreEqual(0, colr.GetTheme());
            //       Assert.IsNotNull( colr.GetRgb() );

            //       // Column 0 has a font with colours from a theme
            //       XSSFCell ct = r.GetCell(0);
            //       XSSFFont ft = wb.GetFontAt( ct.GetCellStyle().GetFontIndex() );
            //       XSSFColor colt =  ft.GetXSSFColor();
            //       // Has a theme, which has the colours on it
            //       Assert.AreEqual(9, colt.GetTheme());
            //       XSSFColor themeC = wb.GetTheme().GetThemeColor(colt.GetTheme());
            //       Assert.IsNotNull( themeC.GetRgb() );
            //       Assert.IsNotNull( colt.GetRgb() );
            //       Assert.AreEqual( themeC.GetARGBHex(), colt.GetARGBHex() ); // The same colour
        }

        /**
         * New lines were being eaten when Setting a font on
         *  a rich text string
         */
        [Test]
        public void Test48877()
        {
            throw new NotImplementedException();
            //       String text = "Use \n with word wrap on to create a new line.\n" +
            //          "This line finishes with two trailing spaces.  ";

            //       XSSFWorkbook wb = new XSSFWorkbook();
            //       XSSFSheet sheet = wb.CreateSheet();

            //       Font font1 = wb.CreateFont();
            //       font1.SetColor((short) 20);
            //       Font font2 = wb.CreateFont();
            //       font2.SetColor(Font.COLOR_RED);
            //       Font font3 = wb.GetFontAt((short)0);

            //       XSSFRow row = sheet.CreateRow(2);
            //       XSSFCell cell = row.CreateCell(2);

            //       XSSFRichTextString richTextString =
            //          wb.GetCreationHelper().CreateRichTextString(text);

            //       // Check the text has the newline
            //       Assert.AreEqual(text, richTextString.GetString());

            //       // Apply the font
            //       richTextString.ApplyFont(font3);
            //       richTextString.ApplyFont(0, 3, font1);
            //       cell.SetCellValue(richTextString);

            //       // To enable newlines you need Set a cell styles with wrap=true
            //       CellStyle cs = wb.CreateCellStyle();
            //       cs.SetWrapText(true);
            //       cell.SetCellStyle(cs);

            //       // Check the text has the
            //       Assert.AreEqual(text, cell.GetStringCellValue());

            //       // Save the file and re-read it
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       sheet = wb.GetSheetAt(0);
            //       row = sheet.GetRow(2);
            //       cell = row.GetCell(2);
            //       Assert.AreEqual(text, cell.GetStringCellValue());

            //       // Now add a 2nd, and check again
            //       int fontAt = text.IndexOf("\n", 6);
            //       cell.GetRichStringCellValue().ApplyFont(10, fontAt+1, font2);
            //       Assert.AreEqual(text, cell.GetStringCellValue());

            //       Assert.AreEqual(4, cell.GetRichStringCellValue().numFormattingRuns());
            //       Assert.AreEqual("Use", cell.GetRichStringCellValue().GetCTRst().GetRList().Get(0).GetT());

            //       String r3 = cell.GetRichStringCellValue().GetCTRst().GetRList().Get(2).GetT();
            //       Assert.AreEqual("line.\n", r3.Substring(r3.Length-6));

            //       // Save and re-check
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       sheet = wb.GetSheetAt(0);
            //       row = sheet.GetRow(2);
            //       cell = row.GetCell(2);
            //       Assert.AreEqual(text, cell.GetStringCellValue());

            ////       FileOutputStream out = new FileOutputStream("/tmp/test48877.xlsx");
            ////       wb.Write(out);
            ////       out.Close();
        }

        /**
         * Adding sheets when one has a table, then re-ordering
         */
        [Test]
        public void Test50867()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50867_with_table.xlsx");
            //       Assert.AreEqual(3, wb.GetNumberOfSheets());

            //       XSSFSheet s1 = wb.GetSheetAt(0);
            //       XSSFSheet s2 = wb.GetSheetAt(1);
            //       XSSFSheet s3 = wb.GetSheetAt(2);
            //       Assert.AreEqual(1, s1.GetTables().Count);
            //       Assert.AreEqual(0, s2.GetTables().Count);
            //       Assert.AreEqual(0, s3.GetTables().Count);

            //       XSSFTable t = s1.GetTables().Get(0);
            //       Assert.AreEqual("Tabella1", t.GetName());
            //       Assert.AreEqual("Tabella1", t.GetDisplayName());
            //       Assert.AreEqual("A1:C3", t.GetCTTable().GetRef());

            //       // Add a sheet and re-order
            //       XSSFSheet s4 = wb.CreateSheet("NewSheet");
            //       wb.SetSheetOrder(s4.GetSheetName(), 0);

            //       // Check on tables
            //       Assert.AreEqual(1, s1.GetTables().Count);
            //       Assert.AreEqual(0, s2.GetTables().Count);
            //       Assert.AreEqual(0, s3.GetTables().Count);
            //       Assert.AreEqual(0, s4.GetTables().Count);

            //       // Refetch to Get the new order
            //       s1 = wb.GetSheetAt(0);
            //       s2 = wb.GetSheetAt(1);
            //       s3 = wb.GetSheetAt(2);
            //       s4 = wb.GetSheetAt(3);
            //       Assert.AreEqual(0, s1.GetTables().Count);
            //       Assert.AreEqual(1, s2.GetTables().Count);
            //       Assert.AreEqual(0, s3.GetTables().Count);
            //       Assert.AreEqual(0, s4.GetTables().Count);

            //       // Save and re-load
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       s1 = wb.GetSheetAt(0);
            //       s2 = wb.GetSheetAt(1);
            //       s3 = wb.GetSheetAt(2);
            //       s4 = wb.GetSheetAt(3);
            //       Assert.AreEqual(0, s1.GetTables().Count);
            //       Assert.AreEqual(1, s2.GetTables().Count);
            //       Assert.AreEqual(0, s3.GetTables().Count);
            //       Assert.AreEqual(0, s4.GetTables().Count);

            //       t = s2.GetTables().Get(0);
            //       Assert.AreEqual("Tabella1", t.GetName());
            //       Assert.AreEqual("Tabella1", t.GetDisplayName());
            //       Assert.AreEqual("A1:C3", t.GetCTTable().GetRef());


            //       // Add some more tables, and check
            //       t = s2.CreateTable();
            //       t.SetName("New 2");
            //       t.SetDisplayName("New 2");
            //       t = s3.CreateTable();
            //       t.SetName("New 3");
            //       t.SetDisplayName("New 3");

            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       s1 = wb.GetSheetAt(0);
            //       s2 = wb.GetSheetAt(1);
            //       s3 = wb.GetSheetAt(2);
            //       s4 = wb.GetSheetAt(3);
            //       Assert.AreEqual(0, s1.GetTables().Count);
            //       Assert.AreEqual(2, s2.GetTables().Count);
            //       Assert.AreEqual(1, s3.GetTables().Count);
            //       Assert.AreEqual(0, s4.GetTables().Count);

            //       t = s2.GetTables().Get(0);
            //       Assert.AreEqual("Tabella1", t.GetName());
            //       Assert.AreEqual("Tabella1", t.GetDisplayName());
            //       Assert.AreEqual("A1:C3", t.GetCTTable().GetRef());

            //       t = s2.GetTables().Get(1);
            //       Assert.AreEqual("New 2", t.GetName());
            //       Assert.AreEqual("New 2", t.GetDisplayName());

            //       t = s3.GetTables().Get(0);
            //       Assert.AreEqual("New 3", t.GetName());
            //       Assert.AreEqual("New 3", t.GetDisplayName());

            //       // Check the relationships
            //       Assert.AreEqual(0, s1.GetRelations().Count);
            //       Assert.AreEqual(3, s2.GetRelations().Count);
            //       Assert.AreEqual(1, s3.GetRelations().Count);
            //       Assert.AreEqual(0, s4.GetRelations().Count);

            //       Assert.AreEqual(
            //             XSSFRelation.PRINTER_SETTINGS.GetContentType(), 
            //             s2.GetRelations().Get(0).GetPackagePart().GetContentType()
            //       );
            //       Assert.AreEqual(
            //             XSSFRelation.TABLE.GetContentType(), 
            //             s2.GetRelations().Get(1).GetPackagePart().GetContentType()
            //       );
            //       Assert.AreEqual(
            //             XSSFRelation.TABLE.GetContentType(), 
            //             s2.GetRelations().Get(2).GetPackagePart().GetContentType()
            //       );
            //       Assert.AreEqual(
            //             XSSFRelation.TABLE.GetContentType(), 
            //             s3.GetRelations().Get(0).GetPackagePart().GetContentType()
            //       );
            //       Assert.AreEqual(
            //             "/xl/tables/table3.xml",
            //             s3.GetRelations().Get(0).GetPackagePart().GetPartName().ToString()
            //       );
        }

        /**
         * Setting repeating rows and columns shouldn't break
         *  any print Settings that were there before
         */
        [Test]
        public void Test49253()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb1 = new XSSFWorkbook();
            //       XSSFWorkbook wb2 = new XSSFWorkbook();

            //       // No print Settings before repeating
            //       XSSFSheet s1 = wb1.CreateSheet(); 
            //       Assert.AreEqual(false, s1.GetCTWorksheet().IsSetPageSetup());
            //       Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());

            //       wb1.SetRepeatingRowsAndColumns(0, 2, 3, 1, 2);

            //       Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageSetup());
            //       Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());

            //       XSSFPrintSetup ps1 = s1.GetPrintSetup();
            //       Assert.AreEqual(false, ps1.GetValidSettings());
            //       Assert.AreEqual(false, ps1.GetLandscape());


            //       // Had valid print Settings before repeating
            //       XSSFSheet s2 = wb2.CreateSheet();
            //       XSSFPrintSetup ps2 = s2.GetPrintSetup();
            //       Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageSetup());
            //       Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageMargins());

            //       ps2.SetLandscape(false);
            //       Assert.AreEqual(true, ps2.GetValidSettings());
            //       Assert.AreEqual(false, ps2.GetLandscape());

            //       wb2.SetRepeatingRowsAndColumns(0, 2, 3, 1, 2);

            //       ps2 = s2.GetPrintSetup();
            //       Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageSetup());
            //       Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageMargins());
            //       Assert.AreEqual(true, ps2.GetValidSettings());
            //       Assert.AreEqual(false, ps2.GetLandscape());
        }

        /**
         * Default Column style
         */
        [Test]
        public void Test51037()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = new XSSFWorkbook();
            //       XSSFSheet s = wb.CreateSheet();

            //       CellStyle defaultStyle = wb.GetCellStyleAt((short)0);
            //       Assert.AreEqual(0, defaultStyle.GetIndex());

            //       CellStyle blueStyle = wb.CreateCellStyle();
            //       blueStyle.SetFillForegroundColor(IndexedColors.AQUA.GetIndex());
            //       blueStyle.SetFillPattern(CellStyle.SOLID_FOREGROUND);
            //       Assert.AreEqual(1, blueStyle.GetIndex());

            //       CellStyle pinkStyle = wb.CreateCellStyle();
            //       pinkStyle.SetFillForegroundColor(IndexedColors.PINK.GetIndex());
            //       pinkStyle.SetFillPattern(CellStyle.SOLID_FOREGROUND);
            //       Assert.AreEqual(2, pinkStyle.GetIndex());

            //       // Starts empty
            //       Assert.AreEqual(1, s.GetCTWorksheet().sizeOfColsArray());
            //       CTCols cols = s.GetCTWorksheet().GetColsArray(0);
            //       Assert.AreEqual(0, cols.sizeOfColArray());

            //       // Add some rows and columns
            //       XSSFRow r1 = s.CreateRow(0);
            //       XSSFRow r2 = s.CreateRow(1);
            //       r1.CreateCell(0);
            //       r1.CreateCell(2);
            //       r2.CreateCell(0);
            //       r2.CreateCell(3);

            //       // Check no style is there
            //       Assert.AreEqual(1, s.GetCTWorksheet().sizeOfColsArray());
            //       Assert.AreEqual(0, cols.sizeOfColArray());

            //       Assert.AreEqual(defaultStyle, s.GetColumnStyle(0));
            //       Assert.AreEqual(defaultStyle, s.GetColumnStyle(2));
            //       Assert.AreEqual(defaultStyle, s.GetColumnStyle(3));


            //       // Apply the styles
            //       s.SetDefaultColumnStyle(0, pinkStyle);
            //       s.SetDefaultColumnStyle(3, blueStyle);

            //       // Check
            //       Assert.AreEqual(pinkStyle, s.GetColumnStyle(0));
            //       Assert.AreEqual(defaultStyle, s.GetColumnStyle(2));
            //       Assert.AreEqual(blueStyle, s.GetColumnStyle(3));

            //       Assert.AreEqual(1, s.GetCTWorksheet().sizeOfColsArray());
            //       Assert.AreEqual(2, cols.sizeOfColArray());

            //       Assert.AreEqual(1, cols.GetColArray(0).GetMin());
            //       Assert.AreEqual(1, cols.GetColArray(0).GetMax());
            //       Assert.AreEqual(pinkStyle.GetIndex(), cols.GetColArray(0).GetStyle());

            //       Assert.AreEqual(4, cols.GetColArray(1).GetMin());
            //       Assert.AreEqual(4, cols.GetColArray(1).GetMax());
            //       Assert.AreEqual(blueStyle.GetIndex(), cols.GetColArray(1).GetStyle());


            //       // Save, re-load and re-check 
            //       wb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       s = wb.GetSheetAt(0);
            //       defaultStyle = wb.GetCellStyleAt(defaultStyle.GetIndex());
            //       blueStyle = wb.GetCellStyleAt(blueStyle.GetIndex());
            //       pinkStyle = wb.GetCellStyleAt(pinkStyle.GetIndex());

            //       Assert.AreEqual(pinkStyle, s.GetColumnStyle(0));
            //       Assert.AreEqual(defaultStyle, s.GetColumnStyle(2));
            //       Assert.AreEqual(blueStyle, s.GetColumnStyle(3));
        }

        /**
         * Repeatedly writing a file.
         * Something with the SharedStringsTable currently breaks...
         */
        public void DISABLEDtest46662()
        {
            throw new NotImplementedException();
            //       // New file
            //       XSSFWorkbook wb = new XSSFWorkbook();
            //       XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       XSSFTestDataSamples.WriteOutAndReadBack(wb);

            //       // Simple file
            //       wb = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            //       XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       XSSFTestDataSamples.WriteOutAndReadBack(wb);
            //       XSSFTestDataSamples.WriteOutAndReadBack(wb);

            //       // Complex file
            //       // TODO
        }

        /**
         * Colours and styles when the list has gaps in it 
         */
        [Test]
        public void Test51222()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51222.xlsx");
            //       XSSFSheet s = wb.GetSheetAt(0);

            //       XSSFCell cA4_EEECE1 = s.GetRow(3).GetCell(0);
            //       XSSFCell cA5_1F497D = s.GetRow(4).GetCell(0);

            //       // Check the text
            //       Assert.AreEqual("A4", cA4_EEECE1.GetRichStringCellValue().GetString());
            //       Assert.AreEqual("A5", cA5_1F497D.GetRichStringCellValue().GetString());

            //       // Check the styles assigned to them
            //       Assert.AreEqual(4, cA4_EEECE1.GetCTCell().GetS());
            //       Assert.AreEqual(5, cA5_1F497D.GetCTCell().GetS());

            //       // Check we look up the correct style
            //       Assert.AreEqual(4, cA4_EEECE1.GetCellStyle().GetIndex());
            //       Assert.AreEqual(5, cA5_1F497D.GetCellStyle().GetIndex());

            //       // Check the Fills on them at the low level
            //       Assert.AreEqual(5, cA4_EEECE1.GetCellStyle().GetCoreXf().GetFillId());
            //       Assert.AreEqual(6, cA5_1F497D.GetCellStyle().GetCoreXf().GetFillId());

            //       // These should reference themes 2 and 3
            //       Assert.AreEqual(2, wb.GetStylesSource().GetFillAt(5).GetCTFill().GetPatternFill().GetFgColor().GetTheme());
            //       Assert.AreEqual(3, wb.GetStylesSource().GetFillAt(6).GetCTFill().GetPatternFill().GetFgColor().GetTheme());

            //       // Ensure we Get the right colours for these themes
            //       // TODO fix
            ////       Assert.AreEqual("FFEEECE1", wb.GetTheme().GetThemeColor(2).GetARGBHex());
            ////       Assert.AreEqual("FF1F497D", wb.GetTheme().GetThemeColor(3).GetARGBHex());

            //       // Finally check the colours on the styles
            //       // TODO fix
            ////       Assert.AreEqual("FFEEECE1", cA4_EEECE1.GetCellStyle().GetFillForegroundXSSFColor().GetARGBHex());
            ////       Assert.AreEqual("FF1F497D", cA5_1F497D.GetCellStyle().GetFillForegroundXSSFColor().GetARGBHex());
        }
        [Test]
        public void Test51470()
        {
            throw new NotImplementedException();
            //        XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51470.xlsx");
            //        XSSFSheet sh0 = wb.GetSheetAt(0);
            //        XSSFSheet sh1 = wb.CloneSheet(0);
            //        List<POIXMLDocumentPart> rels0 = sh0.GetRelations();
            //        List<POIXMLDocumentPart> rels1 = sh1.GetRelations();
            //        Assert.AreEqual(1, rels0.Count);
            //        Assert.AreEqual(1, rels1.Count);

            //        Assert.AreEqual(rels0.Get(0).GetPackageRelationship(), rels1.Get(0).GetPackageRelationship());
        }

        /**
         * Add comments to Sheet 1, when Sheet 2 already has
         *  comments (so /xl/comments1.xml is taken)
         */
        [Test]
        public void Test51850()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51850.xlsx");
            XSSFSheet sh1 = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sh2 = wb.GetSheetAt(1) as XSSFSheet;

            // Sheet 2 has comments
            Assert.IsNotNull(sh2.GetCommentsTable(false));
            Assert.AreEqual(1, sh2.GetCommentsTable(false).GetNumberOfComments());

            // Sheet 1 doesn't (yet)
            Assert.IsNull(sh1.GetCommentsTable(false));

            // Try to add comments to Sheet 1
            ICreationHelper factory = wb.GetCreationHelper();
            IDrawing Drawing = sh1.CreateDrawingPatriarch();

            IClientAnchor anchor = factory.CreateClientAnchor();
            anchor.Col1 = (0);
            anchor.Col2 = (4);
            anchor.Row1 = (0);
            anchor.Row2 = (1);

            IComment comment1 = Drawing.CreateCellComment(anchor);
            comment1.String = (
                  factory.CreateRichTextString("I like this cell. It's my favourite."));
            comment1.Author = ("Bob T. Fish");

            IComment comment2 = Drawing.CreateCellComment(anchor);
            comment2.String = (
                  factory.CreateRichTextString("This is much less fun..."));
            comment2.Author = ("Bob T. Fish");

            ICell c1 = sh1.GetRow(0).CreateCell(4);
            c1.SetCellValue(2.3);
            c1.CellComment = (comment1);

            ICell c2 = sh1.GetRow(0).CreateCell(5);
            c2.SetCellValue(2.1);
            c2.CellComment = (comment2);


            // Save and re-load
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            sh1 = wb.GetSheetAt(0) as XSSFSheet;
            sh2 = wb.GetSheetAt(1) as XSSFSheet;

            // Check the comments
            Assert.IsNotNull(sh2.GetCommentsTable(false));
            Assert.AreEqual(1, sh2.GetCommentsTable(false).GetNumberOfComments());

            Assert.IsNotNull(sh1.GetCommentsTable(false));
            Assert.AreEqual(2, sh1.GetCommentsTable(false).GetNumberOfComments());
        }

        /**
         * Sheet names with a , in them
         */
        [Test]
        public void Test51963()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51963.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            Assert.AreEqual("Abc,1", sheet.SheetName);

            XSSFName name = wb.GetName("Intekon.ProdCodes") as XSSFName;
            Assert.AreEqual("'Abc,1'!$A$1:$A$2", name.RefersToFormula);

            AreaReference ref1 = new AreaReference(name.RefersToFormula);
            Assert.AreEqual(0, ref1.FirstCell.Row);
            Assert.AreEqual(0, ref1.FirstCell.Col);
            Assert.AreEqual(1, ref1.LastCell.Row);
            Assert.AreEqual(0, ref1.LastCell.Col);
        }

        /**
         * Sum across multiple workbooks
         *  eg =SUM($Sheet1.C1:$Sheet4.C1)
         * DISABLED As we can't currently Evaluate these
         */
        public void DISABLEDtest48703()
        {
            throw new NotImplementedException();
            //       XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48703.xlsx");
            //       XSSFSheet sheet = wb.GetSheetAt(0);

            //       // Contains two forms, one with a range and one a list
            //       XSSFRow r1 = sheet.GetRow(0);
            //       XSSFRow r2 = sheet.GetRow(1);
            //       XSSFCell c1 = r1.GetCell(1);
            //       XSSFCell c2 = r2.GetCell(1);

            //       Assert.AreEqual(20.0, c1.GetNumericCellValue());
            //       Assert.AreEqual("SUM(Sheet1!C1,Sheet2!C1,Sheet3!C1,Sheet4!C1)", c1.GetCellFormula());

            //       Assert.AreEqual(20.0, c2.GetNumericCellValue());
            //       Assert.AreEqual("SUM(Sheet1:Sheet4!C1)", c2.GetCellFormula());

            //       // Try Evaluating both
            //       XSSFFormulaEvaluator eval = new XSSFFormulaEvaluator(wb);
            //       Eval.evaluateFormulaCell(c1);
            //       Eval.evaluateFormulaCell(c2);

            //       Assert.AreEqual(20.0, c1.GetNumericCellValue());
            //       Assert.AreEqual(20.0, c2.GetNumericCellValue());
        }

        /**
         * Bugzilla 51710: problems Reading shared formuals from .xlsx
         */
        [Test]
        public void Test51710()
        {
            throw new NotImplementedException();
            //        Workbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51710.xlsx");

            //        String[] columns = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N"};
            //        int rowMax = 500; // bug triggers on row index 59

            //        Sheet sheet = wb.GetSheetAt(0);


            //        // go through all formula cells
            //        for (int rInd = 2; rInd <= rowMax; rInd++) {
            //            Row row = sheet.GetRow(rInd);

            //            for (int cInd = 1; cInd <= 12; cInd++) {
            //                Cell cell = row.GetCell(cInd);
            //                String formula = cell.GetCellFormula();
            //                CellReference ref = new CellReference(cell);

            //                //simulate correct answer
            //                String correct = "$A" + (rInd + 1) + "*" + columns[cInd] + "$2";

            //                Assert.AreEqual("Incorrect formula in " + ref.formatAsString(), correct, formula);
            //            }

            //        }
        }

        /**
         * Bug 53101:
         */
        [Test]
        public void Test5301()
        {
            throw new NotImplementedException();
            //        Workbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("53101.xlsx");
            //        FormulaEvaluator Evaluator =
            //                workbook.GetCreationHelper().CreateFormulaEvaluator();
            //        // A1: SUM(B1: IZ1)
            //        double a1Value =
            //                Evaluator.evaluate(workbook.GetSheetAt(0).GetRow(0).GetCell(0)).GetNumberValue();

            //        // Assert
            //        Assert.AreEqual(259.0, a1Value, 0.0);

            //        // KY: SUM(B1: IZ1)
            //        double ky1Value =
            //                Evaluator.evaluate(workbook.GetSheetAt(0).GetRow(0).GetCell(310)).GetNumberValue();

            //        // Assert
            //        Assert.AreEqual(259.0, a1Value, 0.0);
        }

    }

}