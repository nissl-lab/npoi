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
    using NPOI.OpenXml4Net.OPC;
    using System.Collections.Generic;
    using NPOI.XSSF.UserModel.Extensions;
    using System.IO;
    using NPOI.XSSF.Model;
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
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("45431.xlsm");
            OPCPackage pkg = wb.Package;
            Assert.IsTrue(wb.IsMacroEnabled());

            // Check the various macro related bits can be found
            PackagePart vba = pkg.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/vbaProject.bin")
            );
            Assert.IsNotNull(vba);
            // And the Drawing bit
            PackagePart drw = pkg.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            );
            Assert.IsNotNull(drw);


            // Save and re-open, both still there
            XSSFWorkbook nwb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            OPCPackage nPkg = nwb.Package;
            Assert.IsTrue(nwb.IsMacroEnabled());

            vba = nPkg.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/vbaProject.bin")
            );
            Assert.IsNotNull(vba);
            drw = nPkg.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            );
            Assert.IsNotNull(drw);

            // And again, just to be sure
            nwb = XSSFTestDataSamples.WriteOutAndReadBack(nwb) as XSSFWorkbook;
            nPkg = nwb.Package;
            Assert.IsTrue(nwb.IsMacroEnabled());

            vba = nPkg.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/vbaProject.bin")
            );
            Assert.IsNotNull(vba);
            drw = nPkg.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            );
            Assert.IsNotNull(drw);
        }
        [Test]
        public void Test47504()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("47504.xlsx");
            Assert.AreEqual(1, wb.NumberOfSheets);
            XSSFSheet sh = wb.GetSheetAt(0) as XSSFSheet;
            XSSFDrawing drawing = sh.CreateDrawingPatriarch() as XSSFDrawing;
            List<POIXMLDocumentPart> rels = drawing.GetRelations();
            Assert.AreEqual(1, rels.Count);
            Assert.AreEqual("/xl/drawings/#Sheet1!A1", rels[0].GetPackageRelationship().TargetUri.ToString());

            // And again, just to be sure
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            Assert.AreEqual(1, wb.NumberOfSheets);
            sh = wb.GetSheetAt(0) as XSSFSheet;
            drawing = sh.CreateDrawingPatriarch() as XSSFDrawing;
            rels = drawing.GetRelations();
            Assert.AreEqual(1, rels.Count);
            Assert.AreEqual("/xl/drawings/#Sheet1!A1", rels[0].GetPackageRelationship().TargetUri.ToString());
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
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48923.xlsx");
            Assert.AreEqual(4, wb.NumberOfNames);

            IName b1 = wb.GetName("NameB1");
            IName b2 = wb.GetName("NameB2");
            IName sheet2 = wb.GetName("NameSheet2");
            IName test = wb.GetName("Test");

            Assert.IsNotNull(b1);
            Assert.AreEqual("NameB1", b1.NameName);
            Assert.AreEqual("Sheet1", b1.SheetName);
            Assert.AreEqual(-1, b1.SheetIndex);

            Assert.IsNotNull(b2);
            Assert.AreEqual("NameB2", b2.NameName);
            Assert.AreEqual("Sheet1", b2.SheetName);
            Assert.AreEqual(-1, b2.SheetIndex);

            Assert.IsNotNull(sheet2);
            Assert.AreEqual("NameSheet2", sheet2.NameName);
            Assert.AreEqual("Sheet2", sheet2.SheetName);
            Assert.AreEqual(-1, sheet2.SheetIndex);

            Assert.IsNotNull(test);
            Assert.AreEqual("Test", test.NameName);
            Assert.AreEqual("Sheet1", test.SheetName);
            Assert.AreEqual(-1, test.SheetIndex);
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
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48539.xlsx");
            Assert.AreEqual(3, wb.NumberOfSheets);

            // Try each cell individually
            XSSFFormulaEvaluator eval = new XSSFFormulaEvaluator(wb);
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet s = wb.GetSheetAt(i);
                foreach (IRow r in s)
                {
                    foreach (ICell c in r)
                    {
                        if (c.CellType == CellType.Formula)
                        {
                            CellValue cv = eval.Evaluate(c);
                            if (cv.CellType == CellType.Numeric)
                            {
                                // assert that the calculated value agrees with
                                // the cached formula result calculated by Excel
                                double cachedFormulaResult = c.NumericCellValue;
                                double EvaluatedFormulaResult = cv.NumberValue;
                                Assert.AreEqual(cachedFormulaResult, EvaluatedFormulaResult, 1E-7, c.CellFormula);
                            }
                        }
                    }
                }
            }

            // Now all of them
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
        }

        /**
         * Foreground colours should be found even if
         *  a theme is used 
         */
        [Test]
        public void Test48779()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48779.xlsx");
            XSSFCell cell = wb.GetSheetAt(0).GetRow(0).GetCell(0) as XSSFCell;
            XSSFCellStyle cs = cell.CellStyle as XSSFCellStyle;

            Assert.IsNotNull(cs);
            Assert.AreEqual(1, cs.Index);

            // Look at the low level xml elements
            Assert.AreEqual(2, cs.GetCoreXf().fillId);
            Assert.AreEqual(0, cs.GetCoreXf().xfId);
            Assert.AreEqual(true, cs.GetCoreXf().applyFill);

            XSSFCellFill fg = wb.GetStylesSource().GetFillAt(2);
            Assert.AreEqual(0, fg.GetFillForegroundColor().Indexed);
            Assert.AreEqual(0.0, fg.GetFillForegroundColor().Tint);
            Assert.AreEqual("FFFF0000", fg.GetFillForegroundColor().GetARGBHex());
            Assert.AreEqual(64, fg.GetFillBackgroundColor().Indexed);

            // Now look higher up
            Assert.IsNotNull(cs.FillForegroundXSSFColor);
            Assert.AreEqual(0, cs.FillForegroundColor);
            Assert.AreEqual("FFFF0000", cs.FillForegroundXSSFColor.GetARGBHex());
            Assert.AreEqual("FFFF0000", (cs.FillForegroundColorColor as XSSFColor).GetARGBHex());

            Assert.IsNotNull(cs.FillBackgroundColor);
            Assert.AreEqual(64, cs.FillBackgroundColor);
            Assert.AreEqual(null, cs.FillBackgroundXSSFColor.GetARGBHex());
            Assert.AreEqual(null, (cs.FillBackgroundColorColor as XSSFColor).GetARGBHex());
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
            // Verify the HSSF behaviour
            // Then ensure the same for XSSF
            IWorkbook[] wbs = new IWorkbook[] {
                         new HSSFWorkbook(),
                         new XSSFWorkbook()
                   };
            int[] InitialFonts = new int[] { 4, 1 };
            for (int i = 0; i < wbs.Length; i++)
            {
                IWorkbook wb = wbs[i];
                int startingFonts = InitialFonts[i];

                Assert.AreEqual(startingFonts, wb.NumberOfFonts);

                // Get a font, and slightly change it
                IFont a = wb.CreateFont();
                Assert.AreEqual(startingFonts + 1, wb.NumberOfFonts);
                a.FontHeightInPoints=((short)23);
                Assert.AreEqual(startingFonts + 1, wb.NumberOfFonts);

                // Get two more, unChanged
                IFont b = wb.CreateFont();
                Assert.AreEqual(startingFonts + 2, wb.NumberOfFonts);
                IFont c = wb.CreateFont();
                Assert.AreEqual(startingFonts + 3, wb.NumberOfFonts);
            }
        }

        /**
         * Ensure General and @ format are working properly
         *  for integers 
         */
        [Test]
        public void Test47490()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("GeneralFormatTests.xlsx");
            ISheet s = wb.GetSheetAt(1);
            IRow r;
            DataFormatter df = new DataFormatter();

            r = s.GetRow(1);
            Assert.AreEqual(1.0, r.GetCell(2).NumericCellValue);
            Assert.AreEqual("General", r.GetCell(2).CellStyle.GetDataFormatString());
            Assert.AreEqual("1", df.FormatCellValue(r.GetCell(2)));
            Assert.AreEqual("1", df.FormatRawCellContents(1.0, -1, "@"));
            Assert.AreEqual("1", df.FormatRawCellContents(1.0, -1, "General"));

            r = s.GetRow(2);
            Assert.AreEqual(12.0, r.GetCell(2).NumericCellValue);
            Assert.AreEqual("General", r.GetCell(2).CellStyle.GetDataFormatString());
            Assert.AreEqual("12", df.FormatCellValue(r.GetCell(2)));
            Assert.AreEqual("12", df.FormatRawCellContents(12.0, -1, "@"));
            Assert.AreEqual("12", df.FormatRawCellContents(12.0, -1, "General"));

            r = s.GetRow(3);
            Assert.AreEqual(123.0, r.GetCell(2).NumericCellValue);
            Assert.AreEqual("General", r.GetCell(2).CellStyle.GetDataFormatString());
            Assert.AreEqual("123", df.FormatCellValue(r.GetCell(2)));
            Assert.AreEqual("123", df.FormatRawCellContents(123.0, -1, "@"));
            Assert.AreEqual("123", df.FormatRawCellContents(123.0, -1, "General"));
        }

        /**
         * Ensures that XSSF and HSSF agree with each other,
         *  and with the docs on when fetching the wrong
         *  kind of value from a Formula cell
         */
        [Test]
        public void Test47815()
        {
            IWorkbook[] wbs = new IWorkbook[] {
                         new HSSFWorkbook(),
                         new XSSFWorkbook()
                   };
            foreach (IWorkbook wb in wbs)
            {
                ISheet s = wb.CreateSheet();
                IRow r = s.CreateRow(0);

                // Setup
                ICell cn = r.CreateCell(0, CellType.Numeric);
                cn.SetCellValue(1.2);
                ICell cs = r.CreateCell(1, CellType.String);
                cs.SetCellValue("Testing");

                ICell cfn = r.CreateCell(2, CellType.Formula);
                cfn.SetCellFormula("A1");
                ICell cfs = r.CreateCell(3, CellType.Formula);
                cfs.SetCellFormula("B1");

                IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
                Assert.AreEqual(CellType.Numeric, fe.Evaluate(cfn).CellType);
                Assert.AreEqual(CellType.String, fe.Evaluate(cfs).CellType);
                fe.EvaluateFormulaCell(cfn);
                fe.EvaluateFormulaCell(cfs);

                // Now Test
                Assert.AreEqual(CellType.Numeric, cn.CellType);
                Assert.AreEqual(CellType.String, cs.CellType);
                Assert.AreEqual(CellType.Formula, cfn.CellType);
                Assert.AreEqual(CellType.Numeric, cfn.CachedFormulaResultType);
                Assert.AreEqual(CellType.Formula, cfs.CellType);
                Assert.AreEqual(CellType.String, cfs.CachedFormulaResultType);

                // Different ways of retrieving
                Assert.AreEqual(1.2, cn.NumericCellValue);
                object obj;
                try
                {
                    obj = cn.RichStringCellValue;
                    Assert.Fail();
                }
                catch (InvalidOperationException) { }

                Assert.AreEqual("Testing", cs.StringCellValue);
                try
                {
                    obj = cs.NumericCellValue;
                    Assert.Fail();
                }
                catch (InvalidOperationException) { }

                Assert.AreEqual(1.2, cfn.NumericCellValue);
                try
                {
                    obj = cfn.RichStringCellValue;
                    Assert.Fail();
                }
                catch (InvalidOperationException) { }

                Assert.AreEqual("Testing", cfs.StringCellValue);
                try
                {
                    obj = cfs.NumericCellValue;
                    Assert.Fail();
                }
                catch (InvalidOperationException) { }
            }
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
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("49783.xlsx");
            ISheet sheet = wb.GetSheetAt(0);
            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            ICell cell;

            cell = sheet.GetRow(0).GetCell(0);
            Assert.AreEqual("#REF!*#REF!", cell.CellFormula);
            Assert.AreEqual(CellType.Error, Evaluator.EvaluateInCell(cell).CellType);
            Assert.AreEqual("#REF!", FormulaError.ForInt(cell.ErrorCellValue).String);

            IName nm1 = wb.GetName("sale_1");
            Assert.IsNotNull(nm1, "name sale_1 should be present");
            Assert.AreEqual("Sheet1!#REF!", nm1.RefersToFormula);
            IName nm2 = wb.GetName("sale_2");
            Assert.IsNotNull(nm2, "name sale_2 should be present");
            Assert.AreEqual("Sheet1!#REF!", nm2.RefersToFormula);

            cell = sheet.GetRow(1).GetCell(0);
            Assert.AreEqual("sale_1*sale_2", cell.CellFormula);
            Assert.AreEqual(CellType.Error, Evaluator.EvaluateInCell(cell).CellType);
            Assert.AreEqual("#REF!", FormulaError.ForInt(cell.ErrorCellValue).String);
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
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet s = wb.CreateSheet() as XSSFSheet;
            XSSFRow r = s.CreateRow(0) as XSSFRow;
            XSSFCell c = r.CreateCell(0) as XSSFCell;

            // First without fonts
            c.SetCellValue(
                  new XSSFRichTextString(" with spaces ")
            );
            Assert.AreEqual(" with spaces ", c.RichStringCellValue.ToString());
            Assert.AreEqual(0, (c.RichStringCellValue as XSSFRichTextString).GetCTRst().sizeOfRArray());
            Assert.AreEqual(true, (c.RichStringCellValue as XSSFRichTextString).GetCTRst().IsSetT());
            // Should have the preserve Set
            //Assert.AreEqual(
            //      1,
            //      (c.RichStringCellValue as XSSFRichTextString).GetCTRst().xgetT().GetDomNode().GetAttributes().GetLength()
            //);
            //Assert.AreEqual(
            //      "preserve",
            //      (c.RichStringCellValue as XSSFRichTextString).GetCTRst().xgetT().GetDomNode().GetAttributes().item(0).GetNodeValue()
            //);

            // Save and check
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            s = wb.GetSheetAt(0) as XSSFSheet;
            r = s.GetRow(0) as XSSFRow;
            c = r.GetCell(0) as XSSFCell;
            Assert.AreEqual(" with spaces ", c.RichStringCellValue.ToString());
            Assert.AreEqual(0, (c.RichStringCellValue as XSSFRichTextString).GetCTRst().sizeOfRArray());
            Assert.AreEqual(true, (c.RichStringCellValue as XSSFRichTextString).GetCTRst().IsSetT());

            // Change the string
            c.SetCellValue(
                  new XSSFRichTextString("hello world")
            );
            Assert.AreEqual("hello world", c.RichStringCellValue.ToString());
            // Won't have preserve
            //Assert.AreEqual(
            //      0,
            //      c.RichStringCellValue.GetCTRst().xgetT().GetDomNode().GetAttributes().GetLength()
            //);

            // Apply a font
            XSSFFont f = wb.CreateFont() as XSSFFont;
            f.IsBold=(true);
            c.RichStringCellValue.ApplyFont(0, 5, f);
            Assert.AreEqual("hello world", c.RichStringCellValue.ToString());
            // Does need preserving on the 2nd part
            Assert.AreEqual(2, (c.RichStringCellValue as XSSFRichTextString).GetCTRst().sizeOfRArray());

            //Assert.AreEqual(
            //      0,
            //      c.RichStringCellValue.GetCTRst().GetRArray(0).xgetT().GetDomNode().GetAttributes().GetLength()
            //);
            //Assert.AreEqual(
            //      1,
            //      c.RichStringCellValue.GetCTRst().GetRArray(1).xgetT().GetDomNode().GetAttributes().GetLength()
            //);
            //Assert.AreEqual(
            //      "preserve",
            //      c.RichStringCellValue.GetCTRst().GetRArray(1).xgetT().GetDomNode().GetAttributes().item(0).GetNodeValue()
            //);

            // Save and check
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            s = wb.GetSheetAt(0) as XSSFSheet;
            r = s.GetRow(0) as XSSFRow;
            c = r.GetCell(0) as XSSFCell;
            Assert.AreEqual("hello world", c.RichStringCellValue.ToString());
        }

        /**
         * Repeatedly writing the same file which has styles
         * TODO Currently failing
         */
        public void DISABLEDtest49940()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("styles.xlsx");
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.AreEqual(10, wb.GetStylesSource().GetNumCellStyles());

            MemoryStream b1 = new MemoryStream();
            MemoryStream b2 = new MemoryStream();
            MemoryStream b3 = new MemoryStream();
            wb.Write(b1);
            wb.Write(b2);
            wb.Write(b3);

            foreach (byte[] data in new byte[][] {
                         b1.ToArray(), b2.ToArray(), b3.ToArray()
                   })
            {
                MemoryStream bais = new MemoryStream(data);
                wb = new XSSFWorkbook(bais);
                Assert.AreEqual(3, wb.NumberOfSheets);
                Assert.AreEqual(10, wb.GetStylesSource().GetNumCellStyles());
            }
        }

        /**
         * Various ways of removing a cell formula should all zap
         *  the calcChain entry.
         */
        [Test]
        public void Test49966()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("shared_formulas.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            // CalcChain has lots of entries
            CalculationChain cc = wb.GetCalculationChain();
            Assert.AreEqual("A2", cc.GetCTCalcChain().GetCArray(0).r);
            Assert.AreEqual("A3", cc.GetCTCalcChain().GetCArray(1).r);
            Assert.AreEqual("A4", cc.GetCTCalcChain().GetCArray(2).r);
            Assert.AreEqual("A5", cc.GetCTCalcChain().GetCArray(3).r);
            Assert.AreEqual("A6", cc.GetCTCalcChain().GetCArray(4).r);
            Assert.AreEqual("A7", cc.GetCTCalcChain().GetCArray(5).r);
            Assert.AreEqual("A8", cc.GetCTCalcChain().GetCArray(6).r);
            Assert.AreEqual(40, cc.GetCTCalcChain().SizeOfCArray());

            // Try various ways of changing the formulas
            // If it stays a formula, chain entry should remain
            // Otherwise should go
            sheet.GetRow(1).GetCell(0).SetCellFormula("A1"); // stay
            sheet.GetRow(2).GetCell(0).SetCellFormula(null);  // go
            sheet.GetRow(3).GetCell(0).SetCellType(CellType.Formula); // stay
            sheet.GetRow(4).GetCell(0).SetCellType(CellType.String);  // go
            sheet.GetRow(5).RemoveCell(
                  sheet.GetRow(5).GetCell(0)  // go
            );
            sheet.GetRow(6).GetCell(0).SetCellType(CellType.Blank);  // go
            sheet.GetRow(7).GetCell(0).SetCellValue((String)null);  // go

            // Save and check
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            Assert.AreEqual(35, cc.GetCTCalcChain().SizeOfCArray());

            cc = wb.GetCalculationChain();
            Assert.AreEqual("A2", cc.GetCTCalcChain().GetCArray(0).r);
            Assert.AreEqual("A4", cc.GetCTCalcChain().GetCArray(1).r);
            Assert.AreEqual("A9", cc.GetCTCalcChain().GetCArray(2).r);

        }
        [Test]
        public void Test49156()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("49156.xlsx");
            IFormulaEvaluator formulaEvaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sheet = wb.GetSheetAt(0);
            foreach (IRow row in sheet)
            {
                foreach (ICell cell in row)
                {
                    if (cell.CellType == CellType.Formula)
                    {
                        formulaEvaluator.EvaluateInCell(cell); // caused NPE on some cells
                    }
                }
            }
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
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50795.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            XSSFRow row = sheet.GetRow(0) as XSSFRow;

            XSSFCell cellWith = row.GetCell(0) as XSSFCell;
            XSSFCell cellWithoutComment = row.GetCell(1) as XSSFCell;

            Assert.IsNotNull(cellWith.CellComment);
            Assert.IsNull(cellWithoutComment.CellComment);

            String exp = "\u0410\u0432\u0442\u043e\u0440:\ncomment";
            XSSFComment comment = cellWith.CellComment as XSSFComment;
            Assert.AreEqual(exp, comment.String.String);


            // Check we can write it out and read it back as-is
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            sheet = wb.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(0) as XSSFRow;
            cellWith = row.GetCell(0) as XSSFCell;
            cellWithoutComment = row.GetCell(1) as XSSFCell;

            // Double check things are as expected
            Assert.IsNotNull(cellWith.CellComment);
            Assert.IsNull(cellWithoutComment.CellComment);
            comment = cellWith.CellComment as XSSFComment;
            Assert.AreEqual(exp, comment.String.String);


            // Move the comment
            cellWithoutComment.CellComment=(comment);


            // Write out and re-check
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            sheet = wb.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(0) as XSSFRow;

            // Ensure it swapped over
            cellWith = row.GetCell(0) as XSSFCell;
            cellWithoutComment = row.GetCell(1) as XSSFCell;
            Assert.IsNull(cellWith.CellComment);
            Assert.IsNotNull(cellWithoutComment.CellComment);

            comment = cellWithoutComment.CellComment as XSSFComment;
            Assert.AreEqual(exp, comment.String.String);
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
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50299.xlsx");

            // Check all the colours
            for (int sn = 0; sn < wb.NumberOfSheets; sn++)
            {
                ISheet s1 = wb.GetSheetAt(sn);
                foreach (IRow r in s1)
                {
                    foreach (ICell c in r)
                    {
                        ICellStyle cs = c.CellStyle;
                        if (cs != null)
                        {
                            short xx = cs.FillForegroundColor;
                            Console.WriteLine(xx);
                        }
                    }
                }
            }

            // Check one bit in detail
            // Check that we Get back foreground=0 for the theme colours,
            //  and background=64 for the auto colouring
            ISheet s = wb.GetSheetAt(0);
            Assert.AreEqual(0, s.GetRow(0).GetCell(8).CellStyle.FillForegroundColor);
            Assert.AreEqual(64, s.GetRow(0).GetCell(8).CellStyle.FillBackgroundColor);
            Assert.AreEqual(0, s.GetRow(1).GetCell(8).CellStyle.FillForegroundColor);
            Assert.AreEqual(64, s.GetRow(1).GetCell(8).CellStyle.FillBackgroundColor);
        }

        /**
         * Excel .xls style indexed colours in a .xlsx file
         */
        [Test]
        public void Test50786()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50786-indexed_colours.xlsx");
            XSSFSheet s = wb.GetSheetAt(0) as XSSFSheet;
            XSSFRow r = s.GetRow(2) as XSSFRow;

            // Check we have the right cell
            XSSFCell c = r.GetCell(1) as XSSFCell;
            Assert.AreEqual("test\u00a0", c.RichStringCellValue.String);

            // It should be light green
            XSSFCellStyle cs = c.CellStyle as XSSFCellStyle;
            Assert.AreEqual(42, cs.FillForegroundColor);
            Assert.AreEqual(42, (cs.FillForegroundColorColor as XSSFColor).Indexed);
            Assert.IsNotNull((cs.FillForegroundColorColor as XSSFColor).GetRgb());
            Assert.AreEqual("FFCCFFCC", (cs.FillForegroundColorColor as XSSFColor).GetARGBHex());
        }

        /**
         * If the border colours are Set with themes, then we 
         *  should still be able to Get colours
         */
        [Test]
        public void Test50846()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50846-border_colours.xlsx");

            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;
            XSSFRow row = sheet.GetRow(0) as XSSFRow;

            // Border from a theme, brown
            XSSFCell cellT = row.GetCell(0) as XSSFCell;
            XSSFCellStyle styleT = cellT.CellStyle as XSSFCellStyle;
            XSSFColor colorT = styleT.BottomBorderXSSFColor;

            Assert.AreEqual(5, colorT.GetTheme());
            Assert.AreEqual("FFC0504D", colorT.GetARGBHex());

            // Border from a style direct, red
            XSSFCell cellS = row.GetCell(1) as XSSFCell;
            XSSFCellStyle styleS = cellS.CellStyle as XSSFCellStyle;
            XSSFColor colorS = styleS.BottomBorderXSSFColor;

            Assert.AreEqual(0, colorS.GetTheme());
            Assert.AreEqual("FFFF0000", colorS.GetARGBHex());
        }

        /**
         * Fonts where their colours come from the theme rather
         *  then being Set explicitly still should allow the
         *  fetching of the RGB.
         */
        [Test]
        public void Test50784()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50784-font_theme_colours.xlsx");
            XSSFSheet s = wb.GetSheetAt(0) as XSSFSheet;
            XSSFRow r = s.GetRow(0) as XSSFRow;

            // Column 1 has a font with regular colours
            XSSFCell cr = r.GetCell(1) as XSSFCell;
            XSSFFont fr = wb.GetFontAt(cr.CellStyle.FontIndex) as XSSFFont;
            XSSFColor colr = fr.GetXSSFColor();
            // No theme, has colours
            Assert.AreEqual(0, colr.GetTheme());
            Assert.IsNotNull(colr.GetRgb());

            // Column 0 has a font with colours from a theme
            XSSFCell ct = r.GetCell(0) as XSSFCell;
            XSSFFont ft = wb.GetFontAt(ct.CellStyle.FontIndex) as XSSFFont;
            XSSFColor colt = ft.GetXSSFColor();
            // Has a theme, which has the colours on it
            Assert.AreEqual(9, colt.GetTheme());
            XSSFColor themeC = wb.GetTheme().GetThemeColor(colt.GetTheme());
            Assert.IsNotNull(themeC.GetRgb());
            Assert.IsNotNull(colt.GetRgb());
            Assert.AreEqual(themeC.GetARGBHex(), colt.GetARGBHex()); // The same colour
        }

        /**
         * New lines were being eaten when Setting a font on
         *  a rich text string
         */
        [Test]
        public void Test48877()
        {
            String text = "Use \n with word wrap on to create a new line.\n" +
               "This line finishes with two trailing spaces.  ";

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

            IFont font1 = wb.CreateFont();
            font1.Color=((short)20);
            IFont font2 = wb.CreateFont();
            font2.Color = (short)(FontColor.Red);
            IFont font3 = wb.GetFontAt((short)0);

            XSSFRow row = sheet.CreateRow(2) as XSSFRow;
            XSSFCell cell = row.CreateCell(2) as XSSFCell;

            XSSFRichTextString richTextString =
               wb.GetCreationHelper().CreateRichTextString(text) as XSSFRichTextString;

            // Check the text has the newline
            Assert.AreEqual(text, richTextString.String);

            // Apply the font
            richTextString.ApplyFont(font3);
            richTextString.ApplyFont(0, 3, font1);
            cell.SetCellValue(richTextString);

            // To enable newlines you need Set a cell styles with wrap=true
            ICellStyle cs = wb.CreateCellStyle();
            cs.WrapText=(true);
            cell.CellStyle=(cs);

            // Check the text has the
            Assert.AreEqual(text, cell.StringCellValue);

            // Save the file and re-read it
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            sheet = wb.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(2) as XSSFRow;
            cell = row.GetCell(2) as XSSFCell;
            Assert.AreEqual(text, cell.StringCellValue);

            // Now add a 2nd, and check again
            int fontAt = text.IndexOf("\n", 6);
            cell.RichStringCellValue.ApplyFont(10, fontAt + 1, font2);
            Assert.AreEqual(text, cell.StringCellValue);

            Assert.AreEqual(4, (cell.RichStringCellValue as XSSFRichTextString).NumFormattingRuns);
            Assert.AreEqual("Use", (cell.RichStringCellValue as XSSFRichTextString).GetCTRst().r[0].t);

            String r3 = (cell.RichStringCellValue as XSSFRichTextString).GetCTRst().r[2].t;
            Assert.AreEqual("line.\n", r3.Substring(r3.Length - 6));

            // Save and re-check
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            sheet = wb.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(2) as XSSFRow;
            cell = row.GetCell(2) as XSSFCell;
            Assert.AreEqual(text, cell.StringCellValue);

            //       FileOutputStream out = new FileOutputStream("/tmp/test48877.xlsx");
            //       wb.Write(out);
            //       out.Close();
        }

        /**
         * Adding sheets when one has a table, then re-ordering
         */
        [Test]
        public void Test50867()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50867_with_table.xlsx");
            Assert.AreEqual(3, wb.NumberOfSheets);

            XSSFSheet s1 = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet s2 = wb.GetSheetAt(1) as XSSFSheet;
            XSSFSheet s3 = wb.GetSheetAt(2) as XSSFSheet;
            Assert.AreEqual(1, s1.GetTables().Count);
            Assert.AreEqual(0, s2.GetTables().Count);
            Assert.AreEqual(0, s3.GetTables().Count);

            XSSFTable t = s1.GetTables()[(0)];
            Assert.AreEqual("Tabella1", t.Name);
            Assert.AreEqual("Tabella1", t.DisplayName);
            Assert.AreEqual("A1:C3", t.GetCTTable().@ref);

            // Add a sheet and re-order
            XSSFSheet s4 = wb.CreateSheet("NewSheet") as XSSFSheet;
            wb.SetSheetOrder(s4.SheetName, 0);

            // Check on tables
            Assert.AreEqual(1, s1.GetTables().Count);
            Assert.AreEqual(0, s2.GetTables().Count);
            Assert.AreEqual(0, s3.GetTables().Count);
            Assert.AreEqual(0, s4.GetTables().Count);

            // Refetch to Get the new order
            s1 = wb.GetSheetAt(0) as XSSFSheet;
            s2 = wb.GetSheetAt(1) as XSSFSheet;
            s3 = wb.GetSheetAt(2) as XSSFSheet;
            s4 = wb.GetSheetAt(3) as XSSFSheet;
            Assert.AreEqual(0, s1.GetTables().Count);
            Assert.AreEqual(1, s2.GetTables().Count);
            Assert.AreEqual(0, s3.GetTables().Count);
            Assert.AreEqual(0, s4.GetTables().Count);

            // Save and re-load
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            s1 = wb.GetSheetAt(0) as XSSFSheet;
            s2 = wb.GetSheetAt(1) as XSSFSheet;
            s3 = wb.GetSheetAt(2) as XSSFSheet;
            s4 = wb.GetSheetAt(3) as XSSFSheet;
            Assert.AreEqual(0, s1.GetTables().Count);
            Assert.AreEqual(1, s2.GetTables().Count);
            Assert.AreEqual(0, s3.GetTables().Count);
            Assert.AreEqual(0, s4.GetTables().Count);

            t = s2.GetTables()[(0)];
            Assert.AreEqual("Tabella1", t.Name);
            Assert.AreEqual("Tabella1", t.DisplayName);
            Assert.AreEqual("A1:C3", t.GetCTTable().@ref);


            // Add some more tables, and check
            t = s2.CreateTable();
            t.Name = ("New 2");
            t.DisplayName = ("New 2");
            t = s3.CreateTable();
            t.Name = ("New 3");
            t.DisplayName = ("New 3");

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            s1 = wb.GetSheetAt(0) as XSSFSheet;
            s2 = wb.GetSheetAt(1) as XSSFSheet;
            s3 = wb.GetSheetAt(2) as XSSFSheet;
            s4 = wb.GetSheetAt(3) as XSSFSheet;
            Assert.AreEqual(0, s1.GetTables().Count);
            Assert.AreEqual(2, s2.GetTables().Count);
            Assert.AreEqual(1, s3.GetTables().Count);
            Assert.AreEqual(0, s4.GetTables().Count);

            t = s2.GetTables()[(0)];
            Assert.AreEqual("Tabella1", t.Name);
            Assert.AreEqual("Tabella1", t.DisplayName);
            Assert.AreEqual("A1:C3", t.GetCTTable().@ref);

            t = s2.GetTables()[(1)];
            Assert.AreEqual("New 2", t.Name);
            Assert.AreEqual("New 2", t.DisplayName);

            t = s3.GetTables()[(0)];
            Assert.AreEqual("New 3", t.Name);
            Assert.AreEqual("New 3", t.DisplayName);

            // Check the relationships
            Assert.AreEqual(0, s1.GetRelations().Count);
            Assert.AreEqual(3, s2.GetRelations().Count);
            Assert.AreEqual(1, s3.GetRelations().Count);
            Assert.AreEqual(0, s4.GetRelations().Count);

            Assert.AreEqual(
                  XSSFRelation.PRINTER_SETTINGS.ContentType,
                  s2.GetRelations()[0].GetPackagePart().ContentType
            );
            Assert.AreEqual(
                  XSSFRelation.TABLE.ContentType,
                  s2.GetRelations()[1].GetPackagePart().ContentType
            );
            Assert.AreEqual(
                  XSSFRelation.TABLE.ContentType,
                  s2.GetRelations()[2].GetPackagePart().ContentType
            );
            Assert.AreEqual(
                  XSSFRelation.TABLE.ContentType,
                  s3.GetRelations()[(0)].GetPackagePart().ContentType
            );
            Assert.AreEqual(
                  "/xl/tables/table3.xml",
                  s3.GetRelations()[(0)].GetPackagePart().PartName.ToString()
            );
        }

        /**
         * Setting repeating rows and columns shouldn't break
         *  any print Settings that were there before
         */
        [Test]
        public void Test49253()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFWorkbook wb2 = new XSSFWorkbook();

            // No print Settings before repeating
            XSSFSheet s1 = wb1.CreateSheet() as XSSFSheet;
            Assert.AreEqual(false, s1.GetCTWorksheet().IsSetPageSetup());
            Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());

            wb1.SetRepeatingRowsAndColumns(0, 2, 3, 1, 2);

            Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageSetup());
            Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());

            XSSFPrintSetup ps1 = s1.PrintSetup as XSSFPrintSetup;
            Assert.AreEqual(false, ps1.ValidSettings);
            Assert.AreEqual(false, ps1.Landscape);


            // Had valid print Settings before repeating
            XSSFSheet s2 = wb2.CreateSheet() as XSSFSheet;
            XSSFPrintSetup ps2 = s2.PrintSetup as XSSFPrintSetup;
            Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageSetup());
            Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageMargins());

            ps2.Landscape=(false);
            Assert.AreEqual(true, ps2.ValidSettings);
            Assert.AreEqual(false, ps2.Landscape);

            wb2.SetRepeatingRowsAndColumns(0, 2, 3, 1, 2);

            ps2 = s2.PrintSetup as XSSFPrintSetup;
            Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageSetup());
            Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageMargins());
            Assert.AreEqual(true, ps2.ValidSettings);
            Assert.AreEqual(false, ps2.Landscape);
        }

        /**
         * Default Column style
         */
        [Test]
        public void Test51037()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet s = wb.CreateSheet() as XSSFSheet;

            ICellStyle defaultStyle = wb.GetCellStyleAt((short)0);
            Assert.AreEqual(0, defaultStyle.Index);

            ICellStyle blueStyle = wb.CreateCellStyle();
            blueStyle.FillForegroundColor=(IndexedColors.Aqua.Index);
            blueStyle.FillPattern=(FillPattern.SolidForeground);
            Assert.AreEqual(1, blueStyle.Index);

            ICellStyle pinkStyle = wb.CreateCellStyle();
            pinkStyle.FillForegroundColor=(IndexedColors.Pink.Index);
            pinkStyle.FillPattern=(FillPattern.SolidForeground);
            Assert.AreEqual(2, pinkStyle.Index);

            // Starts empty
            Assert.AreEqual(1, s.GetCTWorksheet().sizeOfColsArray());
            CT_Cols cols = s.GetCTWorksheet().GetColsArray(0);
            Assert.AreEqual(0, cols.sizeOfColArray());

            // Add some rows and columns
            XSSFRow r1 = s.CreateRow(0) as XSSFRow;
            XSSFRow r2 = s.CreateRow(1) as XSSFRow;
            r1.CreateCell(0);
            r1.CreateCell(2);
            r2.CreateCell(0);
            r2.CreateCell(3);

            // Check no style is there
            Assert.AreEqual(1, s.GetCTWorksheet().sizeOfColsArray());
            Assert.AreEqual(0, cols.sizeOfColArray());

            Assert.AreEqual(defaultStyle, s.GetColumnStyle(0));
            Assert.AreEqual(defaultStyle, s.GetColumnStyle(2));
            Assert.AreEqual(defaultStyle, s.GetColumnStyle(3));


            // Apply the styles
            s.SetDefaultColumnStyle(0, pinkStyle);
            s.SetDefaultColumnStyle(3, blueStyle);

            // Check
            Assert.AreEqual(pinkStyle, s.GetColumnStyle(0));
            Assert.AreEqual(defaultStyle, s.GetColumnStyle(2));
            Assert.AreEqual(blueStyle, s.GetColumnStyle(3));

            Assert.AreEqual(1, s.GetCTWorksheet().sizeOfColsArray());
            Assert.AreEqual(2, cols.sizeOfColArray());

            Assert.AreEqual(1, cols.GetColArray(0).min);
            Assert.AreEqual(1, cols.GetColArray(0).max);
            Assert.AreEqual(pinkStyle.Index, cols.GetColArray(0).style);

            Assert.AreEqual(4, cols.GetColArray(1).min);
            Assert.AreEqual(4, cols.GetColArray(1).max);
            Assert.AreEqual(blueStyle.Index, cols.GetColArray(1).style);


            // Save, re-load and re-check 
            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            s = wb.GetSheetAt(0) as XSSFSheet;
            defaultStyle = wb.GetCellStyleAt(defaultStyle.Index);
            blueStyle = wb.GetCellStyleAt(blueStyle.Index);
            pinkStyle = wb.GetCellStyleAt(pinkStyle.Index);

            Assert.AreEqual(pinkStyle, s.GetColumnStyle(0));
            Assert.AreEqual(defaultStyle, s.GetColumnStyle(2));
            Assert.AreEqual(blueStyle, s.GetColumnStyle(3));
        }

        /**
         * Repeatedly writing a file.
         * Something with the SharedStringsTable currently breaks...
         */
        public void DISABLEDtest46662()
        {
            // New file
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFTestDataSamples.WriteOutAndReadBack(wb);
            XSSFTestDataSamples.WriteOutAndReadBack(wb);
            XSSFTestDataSamples.WriteOutAndReadBack(wb);

            // Simple file
            wb = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            XSSFTestDataSamples.WriteOutAndReadBack(wb);
            XSSFTestDataSamples.WriteOutAndReadBack(wb);
            XSSFTestDataSamples.WriteOutAndReadBack(wb);

            // Complex file
            // TODO
        }

        /**
         * Colours and styles when the list has gaps in it 
         */
        [Test]
        public void Test51222()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51222.xlsx");
            XSSFSheet s = wb.GetSheetAt(0) as XSSFSheet;

            XSSFCell cA4_EEECE1 = s.GetRow(3).GetCell(0) as XSSFCell;
            XSSFCell cA5_1F497D = s.GetRow(4).GetCell(0) as XSSFCell;

            // Check the text
            Assert.AreEqual("A4", cA4_EEECE1.RichStringCellValue.String);
            Assert.AreEqual("A5", cA5_1F497D.RichStringCellValue.String);

            // Check the styles assigned to them
            Assert.AreEqual(4, cA4_EEECE1.GetCTCell().s);
            Assert.AreEqual(5, cA5_1F497D.GetCTCell().s);

            // Check we look up the correct style
            Assert.AreEqual(4, cA4_EEECE1.CellStyle.Index);
            Assert.AreEqual(5, cA5_1F497D.CellStyle.Index);

            // Check the Fills on them at the low level
            Assert.AreEqual(5, (cA4_EEECE1.CellStyle as XSSFCellStyle).GetCoreXf().fillId);
            Assert.AreEqual(6, (cA5_1F497D.CellStyle as XSSFCellStyle).GetCoreXf().fillId);

            // These should reference themes 2 and 3
            Assert.AreEqual(2, wb.GetStylesSource().GetFillAt(5).GetCTFill().GetPatternFill().fgColor.theme);
            Assert.AreEqual(3, wb.GetStylesSource().GetFillAt(6).GetCTFill().GetPatternFill().fgColor.theme);

            // Ensure we Get the right colours for these themes
            // TODO fix
            //       Assert.AreEqual("FFEEECE1", wb.GetTheme().GetThemeColor(2).GetARGBHex());
            //       Assert.AreEqual("FF1F497D", wb.GetTheme().GetThemeColor(3).GetARGBHex());

            // Finally check the colours on the styles
            // TODO fix
            //       Assert.AreEqual("FFEEECE1", cA4_EEECE1.CellStyle.GetFillForegroundXSSFColor().GetARGBHex());
            //       Assert.AreEqual("FF1F497D", cA5_1F497D.CellStyle.GetFillForegroundXSSFColor().GetARGBHex());
        }
        [Test]
        public void Test51470()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51470.xlsx");
            XSSFSheet sh0 = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sh1 = wb.CloneSheet(0) as XSSFSheet;
            List<POIXMLDocumentPart> rels0 = sh0.GetRelations();
            List<POIXMLDocumentPart> rels1 = sh1.GetRelations();
            Assert.AreEqual(1, rels0.Count);
            Assert.AreEqual(1, rels1.Count);

            Assert.AreEqual(rels0[(0)].GetPackageRelationship(), rels1[0].GetPackageRelationship());
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
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48703.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            // Contains two forms, one with a range and one a list
            XSSFRow r1 = sheet.GetRow(0) as XSSFRow;
            XSSFRow r2 = sheet.GetRow(1) as XSSFRow;
            XSSFCell c1 = r1.GetCell(1) as XSSFCell;
            XSSFCell c2 = r2.GetCell(1) as XSSFCell;

            Assert.AreEqual(20.0, c1.NumericCellValue);
            Assert.AreEqual("SUM(Sheet1!C1,Sheet2!C1,Sheet3!C1,Sheet4!C1)", c1.CellFormula);

            Assert.AreEqual(20.0, c2.NumericCellValue);
            Assert.AreEqual("SUM(Sheet1:Sheet4!C1)", c2.CellFormula);

            // Try Evaluating both
            XSSFFormulaEvaluator eval = new XSSFFormulaEvaluator(wb);
            eval.EvaluateFormulaCell(c1);
            eval.EvaluateFormulaCell(c2);

            Assert.AreEqual(20.0, c1.NumericCellValue);
            Assert.AreEqual(20.0, c2.NumericCellValue);
        }

        /**
         * Bugzilla 51710: problems Reading shared formuals from .xlsx
         */
        [Test]
        public void Test51710()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51710.xlsx");

            String[] columns = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" };
            int rowMax = 500; // bug triggers on row index 59

            ISheet sheet = wb.GetSheetAt(0);


            // go through all formula cells
            for (int rInd = 2; rInd <= rowMax; rInd++)
            {
                IRow row = sheet.GetRow(rInd);

                for (int cInd = 1; cInd <= 12; cInd++)
                {
                    ICell cell = row.GetCell(cInd);
                    String formula = cell.CellFormula;
                    CellReference ref1 = new CellReference(cell);

                    //simulate correct answer
                    String correct = "$A" + (rInd + 1) + "*" + columns[cInd] + "$2";

                    Assert.AreEqual(correct, formula, "Incorrect formula in " + ref1.FormatAsString());
                }

            }
        }

        /**
         * Bug 53101:
         */
        [Test]
        public void Test5301()
        {
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("53101.xlsx");
            IFormulaEvaluator Evaluator =
                    workbook.GetCreationHelper().CreateFormulaEvaluator();
            // A1: SUM(B1: IZ1)
            double a1Value =
                    Evaluator.Evaluate(workbook.GetSheetAt(0).GetRow(0).GetCell(0)).NumberValue;

            // Assert
            Assert.AreEqual(259.0, a1Value, 0.0);

            // KY: SUM(B1: IZ1)
            double ky1Value =
                    Evaluator.Evaluate(workbook.GetSheetAt(0).GetRow(0).GetCell(310)).NumberValue;

            // Assert
            Assert.AreEqual(259.0, a1Value, 0.0);
        }

    }

}