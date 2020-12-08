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
namespace TestCases.XSSF.UserModel
{
    using NPOI;
    using NPOI.HSSF.UserModel;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.Streaming;
    using NPOI.XSSF.UserModel;
    using NPOI.XSSF.UserModel.Extensions;
    using NUnit.Framework;
    using NUnit.Framework.Constraints;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Text;
    using TestCases;
    using TestCases.HSSF;

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

            XSSFName name = wb.GetName("SheetAA1") as XSSFName;
            Assert.AreEqual(0, name.GetCTName().localSheetId);
            Assert.IsFalse(name.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual("SheetA!$A$1", name.RefersToFormula);
            Assert.AreEqual("SheetA", name.SheetName);

            name = wb.GetName("SheetBA1") as XSSFName;
            Assert.AreEqual(0, name.GetCTName().localSheetId);
            Assert.IsFalse(name.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual("SheetB!$A$1", name.RefersToFormula);
            Assert.AreEqual("SheetB", name.SheetName);

            name = wb.GetName("SheetCA1") as XSSFName;
            Assert.AreEqual(0, name.GetCTName().localSheetId);
            Assert.IsFalse(name.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual("SheetC!$A$1", name.RefersToFormula);
            Assert.AreEqual("SheetC", name.SheetName);

            // Save and re-load, still there
            XSSFWorkbook nwb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            Assert.AreEqual(3, nwb.NumberOfNames);
            Assert.AreEqual("SheetA!$A$1", nwb.GetName("SheetAA1").RefersToFormula);

            nwb.Close();
            wb.Close();
        }

        /**
         * We should carry vba macros over After save
         */
        [Test]
        public void Test45431()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("45431.xlsm");
            OPCPackage pkg1 = wb1.Package;
            Assert.IsTrue(wb1.IsMacroEnabled());

            // Check the various macro related bits can be found
            PackagePart vba = pkg1.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/vbaProject.bin")
            );
            Assert.IsNotNull(vba);
            // And the Drawing bit
            PackagePart drw = pkg1.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            );
            Assert.IsNotNull(drw);


            // Save and re-open, both still there
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            pkg1.Close();
            wb1.Close();

            OPCPackage pkg2 = wb2.Package;
            Assert.IsTrue(wb2.IsMacroEnabled());

            vba = pkg2.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/vbaProject.bin")
            );
            Assert.IsNotNull(vba);
            drw = pkg2.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            );
            Assert.IsNotNull(drw);

            // And again, just to be sure
            XSSFWorkbook wb3 = XSSFTestDataSamples.WriteOutAndReadBack(wb2) as XSSFWorkbook;
            pkg2.Close();
            wb2.Close();

            OPCPackage pkg3 = wb3.Package;
            Assert.IsTrue(wb3.IsMacroEnabled());

            vba = pkg3.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/vbaProject.bin")
            );
            Assert.IsNotNull(vba);
            drw = pkg3.GetPart(
                    PackagingUriHelper.CreatePartName("/xl/drawings/vmlDrawing1.vml")
            );
            Assert.IsNotNull(drw);

            pkg3.Close();
            wb3.Close();
        }
        [Test]
        public void Test47504()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("47504.xlsx");
            Assert.AreEqual(1, wb1.NumberOfSheets);
            XSSFSheet sh = wb1.GetSheetAt(0) as XSSFSheet;
            XSSFDrawing drawing = sh.CreateDrawingPatriarch() as XSSFDrawing;
            List<POIXMLDocumentPart.RelationPart> rels = drawing.RelationParts;
            Assert.AreEqual(1, rels.Count);
            Uri baseUri = new Uri("ooxml://npoi.org"); //For test only.
            Uri target = new Uri(baseUri, rels[0].Relationship.TargetUri.ToString());
            Assert.AreEqual("#Sheet1!A1", target.Fragment);

            // And again, just to be sure
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();

            Assert.AreEqual(1, wb2.NumberOfSheets);
            sh = wb2.GetSheetAt(0) as XSSFSheet;
            drawing = sh.CreateDrawingPatriarch() as XSSFDrawing;
            rels = drawing.RelationParts;
            Assert.AreEqual(1, rels.Count);
            Uri target2 = new Uri(baseUri, rels[0].Relationship.TargetUri.ToString());
            Assert.AreEqual("#Sheet1!A1", target2.Fragment);

            wb2.Close();
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
            wb.Close();
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

            wb.Close();
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
        public void Bug48539()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48539.xlsx");
            try
            {
                Assert.AreEqual(3, wb.NumberOfSheets);
                Assert.AreEqual(0, wb.NumberOfNames);

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
            finally
            {
                wb.Close();
            }
            
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
            Assert.IsNotNull(fg.GetFillForegroundColor());
            Assert.AreEqual(0, fg.GetFillForegroundColor().Indexed);
            Assert.AreEqual(0.0, fg.GetFillForegroundColor().Tint);
            Assert.AreEqual("FFFF0000", fg.GetFillForegroundColor().ARGBHex);
            Assert.IsNotNull(fg.GetFillBackgroundColor());
            Assert.AreEqual(64, fg.GetFillBackgroundColor().Indexed);

            // Now look higher up
            Assert.IsNotNull(cs.FillForegroundXSSFColor);
            Assert.AreEqual(0, cs.FillForegroundColor);
            Assert.AreEqual("FFFF0000", cs.FillForegroundXSSFColor.ARGBHex);
            Assert.AreEqual("FFFF0000", (cs.FillForegroundColorColor as XSSFColor).ARGBHex);

            Assert.IsNotNull(cs.FillBackgroundColor);
            Assert.AreEqual(64, cs.FillBackgroundColor);
            Assert.AreEqual(null, cs.FillBackgroundXSSFColor.ARGBHex);
            Assert.AreEqual(null, (cs.FillBackgroundColorColor as XSSFColor).ARGBHex);

            wb.Close();
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
                a.FontHeightInPoints = ((short)23);
                Assert.AreEqual(startingFonts + 1, wb.NumberOfFonts);

                // Get two more, unChanged
                /*IFont b = */
                wb.CreateFont();
                Assert.AreEqual(startingFonts + 2, wb.NumberOfFonts);
                /*IFont c = */
                wb.CreateFont();
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

            wb.Close();
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

            wb.Close();

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

            wb.Close();
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
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet s = wb1.CreateSheet() as XSSFSheet;
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
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();
            s = wb2.GetSheetAt(0) as XSSFSheet;
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
            XSSFFont f = wb2.CreateFont() as XSSFFont;
            f.IsBold = (true);
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
            XSSFWorkbook wb3 = XSSFTestDataSamples.WriteOutAndReadBack(wb2) as XSSFWorkbook;
            wb2.Close();
            s = wb3.GetSheetAt(0) as XSSFSheet;
            r = s.GetRow(0) as XSSFRow;
            c = r.GetCell(0) as XSSFCell;
            Assert.AreEqual("hello world", c.RichStringCellValue.ToString());

            wb3.Close();
        }

        /**
         * Repeatedly writing the same file which has styles
         */
        [Test]
        public void Test49940()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("styles.xlsx");
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.AreEqual(10, wb.GetStylesSource().NumCellStyles);

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
                XSSFWorkbook wb2 = new XSSFWorkbook(bais);
                Assert.AreEqual(3, wb2.NumberOfSheets);
                Assert.AreEqual(10, wb2.GetStylesSource().NumCellStyles);
                wb2.Close();
            }
            wb.Close();
        }

        /**
         * Various ways of removing a cell formula should all zap
         *  the calcChain entry.
         */
        [Test]
        public void Test49966()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("shared_formulas.xlsx");
            XSSFSheet sheet = wb1.GetSheetAt(0) as XSSFSheet;

            XSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();

            // CalcChain has lots of entries
            CalculationChain cc = wb1.GetCalculationChain();
            Assert.AreEqual("A2", cc.GetCTCalcChain().GetCArray(0).r);
            Assert.AreEqual("A3", cc.GetCTCalcChain().GetCArray(1).r);
            Assert.AreEqual("A4", cc.GetCTCalcChain().GetCArray(2).r);
            Assert.AreEqual("A5", cc.GetCTCalcChain().GetCArray(3).r);
            Assert.AreEqual("A6", cc.GetCTCalcChain().GetCArray(4).r);
            Assert.AreEqual("A7", cc.GetCTCalcChain().GetCArray(5).r);
            Assert.AreEqual("A8", cc.GetCTCalcChain().GetCArray(6).r);
            Assert.AreEqual(40, cc.GetCTCalcChain().SizeOfCArray());

            XSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();

            // Try various ways of changing the formulas
            // If it stays a formula, chain entry should remain
            // Otherwise should go
            sheet.GetRow(1).GetCell(0).SetCellFormula("A1"); // stay
            sheet.GetRow(2).GetCell(0).SetCellFormula(null);  // go
            sheet.GetRow(3).GetCell(0).SetCellType(CellType.Formula); // stay
            
            XSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();
            sheet.GetRow(4).GetCell(0).SetCellType(CellType.String);  // go
            
            XSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();

            validateCells(sheet);
            sheet.GetRow(5).RemoveCell(
                  sheet.GetRow(5).GetCell(0)  // go
            );
            validateCells(sheet);
            
            XSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();

            sheet.GetRow(6).GetCell(0).SetCellType(CellType.Blank);  // go
            
            XSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();
            sheet.GetRow(7).GetCell(0).SetCellValue((String)null);  // go
            
            XSSFTestDataSamples.WriteOutAndReadBack(wb1).Close();

            // Save and check
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();
            Assert.AreEqual(35, cc.GetCTCalcChain().SizeOfCArray());

            cc = wb2.GetCalculationChain();
            Assert.AreEqual("A2", cc.GetCTCalcChain().GetCArray(0).r);
            Assert.AreEqual("A4", cc.GetCTCalcChain().GetCArray(1).r);
            Assert.AreEqual("A9", cc.GetCTCalcChain().GetCArray(2).r);
            wb2.Close();
        }
        [Test]
        public void Bug49966Row()
        {
            XSSFWorkbook wb = XSSFTestDataSamples
                    .OpenSampleWorkbook("shared_formulas.xlsx");
            XSSFSheet sheet = wb.GetSheetAt(0) as XSSFSheet;

            validateCells(sheet);
            sheet.GetRow(5).RemoveCell(sheet.GetRow(5).GetCell(0)); // go
            validateCells(sheet);

            wb.Close();
        }

        private void validateCells(XSSFSheet sheet)
        {
            foreach (IRow row in sheet)
            {
                // trigger handling
                ((XSSFRow)row).OnDocumentWrite();
            }
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
            wb.Close();
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

            wb.Close();
        }

        /**
         * Moving a cell comment from one cell to another
         */
        [Test]
        public void Test50795()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("50795.xlsx");
            XSSFSheet sheet = wb1.GetSheetAt(0) as XSSFSheet;
            XSSFRow row = sheet.GetRow(0) as XSSFRow;

            XSSFCell cellWith = row.GetCell(0) as XSSFCell;
            XSSFCell cellWithoutComment = row.GetCell(1) as XSSFCell;

            Assert.IsNotNull(cellWith.CellComment);
            Assert.IsNull(cellWithoutComment.CellComment);

            String exp = "\u0410\u0432\u0442\u043e\u0440:\ncomment";
            XSSFComment comment = cellWith.CellComment as XSSFComment;
            Assert.AreEqual(exp, comment.String.String);


            // Check we can write it out and read it back as-is
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();
            sheet = wb2.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(0) as XSSFRow;
            cellWith = row.GetCell(0) as XSSFCell;
            cellWithoutComment = row.GetCell(1) as XSSFCell;

            // Double check things are as expected
            Assert.IsNotNull(cellWith.CellComment);
            Assert.IsNull(cellWithoutComment.CellComment);
            comment = cellWith.CellComment as XSSFComment;
            Assert.AreEqual(exp, comment.String.String);


            // Move the comment
            cellWithoutComment.CellComment = (comment);


            // Write out and re-check
            XSSFWorkbook wb3 = XSSFTestDataSamples.WriteOutAndReadBack(wb2) as XSSFWorkbook;
            wb2.Close();

            sheet = wb3.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(0) as XSSFRow;

            // Ensure it swapped over
            cellWith = row.GetCell(0) as XSSFCell;
            cellWithoutComment = row.GetCell(1) as XSSFCell;
            Assert.IsNull(cellWith.CellComment);
            Assert.IsNotNull(cellWithoutComment.CellComment);

            comment = cellWithoutComment.CellComment as XSSFComment;
            Assert.AreEqual(exp, comment.String.String);

            wb3.Close();
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

            wb.Close();
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
            Assert.IsNotNull((cs.FillForegroundColorColor as XSSFColor).RGB);
            Assert.AreEqual("FFCCFFCC", (cs.FillForegroundColorColor as XSSFColor).ARGBHex);

            wb.Close();
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

            Assert.AreEqual(5, colorT.Theme);
            Assert.AreEqual("FFC0504D", colorT.ARGBHex);

            // Border from a style direct, red
            XSSFCell cellS = row.GetCell(1) as XSSFCell;
            XSSFCellStyle styleS = cellS.CellStyle as XSSFCellStyle;
            XSSFColor colorS = styleS.BottomBorderXSSFColor;

            Assert.AreEqual(0, colorS.Theme);
            Assert.AreEqual("FFFF0000", colorS.ARGBHex);

            wb.Close();
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
            XSSFColor colr = fr.GetXSSFColor();            // No theme, has colours
            Assert.AreEqual(0, colr.Theme);
            Assert.IsNotNull(colr.GetRgb());

            // Column 0 has a font with colours from a theme
            XSSFCell ct = r.GetCell(0) as XSSFCell;
            XSSFFont ft = wb.GetFontAt(ct.CellStyle.FontIndex) as XSSFFont;
            XSSFColor colt = ft.GetXSSFColor();
            // Has a theme, which has the colours on it
            Assert.AreEqual(9, colt.Theme);
            XSSFColor themeC = wb.GetTheme().GetThemeColor(colt.Theme);
            Assert.IsNotNull(themeC.GetRgb());
            Assert.IsNotNull(colt.GetRgb());
            Assert.AreEqual(themeC.ARGBHex, colt.ARGBHex); // The same colour

            wb.Close();
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

            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet = wb1.CreateSheet() as XSSFSheet;

            IFont font1 = wb1.CreateFont();
            font1.Color = ((short)20);
            IFont font2 = wb1.CreateFont();
            font2.Color = (short)(FontColor.Red);
            IFont font3 = wb1.GetFontAt((short)0);

            XSSFRow row = sheet.CreateRow(2) as XSSFRow;
            XSSFCell cell = row.CreateCell(2) as XSSFCell;

            XSSFRichTextString richTextString =
               wb1.GetCreationHelper().CreateRichTextString(text) as XSSFRichTextString;

            // Check the text has the newline
            Assert.AreEqual(text, richTextString.String);

            // Apply the font
            richTextString.ApplyFont(font3);
            richTextString.ApplyFont(0, 3, font1);
            cell.SetCellValue(richTextString);

            // To enable newlines you need Set a cell styles with wrap=true
            ICellStyle cs = wb1.CreateCellStyle();
            cs.WrapText = (true);
            cell.CellStyle = (cs);

            // Check the text has the
            Assert.AreEqual(text, cell.StringCellValue);

            // Save the file and re-read it
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();

            sheet = wb2.GetSheetAt(0) as XSSFSheet;
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
            XSSFWorkbook wb3 = XSSFTestDataSamples.WriteOutAndReadBack(wb2) as XSSFWorkbook;
            wb2.Close();

            sheet = wb3.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(2) as XSSFRow;
            cell = row.GetCell(2) as XSSFCell;
            Assert.AreEqual(text, cell.StringCellValue);
            wb3.Close();

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
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("50867_with_table.xlsx");
            Assert.AreEqual(3, wb1.NumberOfSheets);

            XSSFSheet s1 = wb1.GetSheetAt(0) as XSSFSheet;
            XSSFSheet s2 = wb1.GetSheetAt(1) as XSSFSheet;
            XSSFSheet s3 = wb1.GetSheetAt(2) as XSSFSheet;
            Assert.AreEqual(1, s1.GetTables().Count);
            Assert.AreEqual(0, s2.GetTables().Count);
            Assert.AreEqual(0, s3.GetTables().Count);

            XSSFTable t = s1.GetTables()[(0)];
            Assert.AreEqual("Tabella1", t.Name);
            Assert.AreEqual("Tabella1", t.DisplayName);
            Assert.AreEqual("A1:C3", t.GetCTTable().@ref);

            // Add a sheet and re-order
            XSSFSheet s4 = wb1.CreateSheet("NewSheet") as XSSFSheet;
            wb1.SetSheetOrder(s4.SheetName, 0);

            // Check on tables
            Assert.AreEqual(1, s1.GetTables().Count);
            Assert.AreEqual(0, s2.GetTables().Count);
            Assert.AreEqual(0, s3.GetTables().Count);
            Assert.AreEqual(0, s4.GetTables().Count);

            // Refetch to Get the new order
            s1 = wb1.GetSheetAt(0) as XSSFSheet;
            s2 = wb1.GetSheetAt(1) as XSSFSheet;
            s3 = wb1.GetSheetAt(2) as XSSFSheet;
            s4 = wb1.GetSheetAt(3) as XSSFSheet;
            Assert.AreEqual(0, s1.GetTables().Count);
            Assert.AreEqual(1, s2.GetTables().Count);
            Assert.AreEqual(0, s3.GetTables().Count);
            Assert.AreEqual(0, s4.GetTables().Count);

            // Save and re-load
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();

            s1 = wb2.GetSheetAt(0) as XSSFSheet;
            s2 = wb2.GetSheetAt(1) as XSSFSheet;
            s3 = wb2.GetSheetAt(2) as XSSFSheet;
            s4 = wb2.GetSheetAt(3) as XSSFSheet;
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

            XSSFWorkbook wb3 = XSSFTestDataSamples.WriteOutAndReadBack(wb2) as XSSFWorkbook;
            wb2.Close();

            s1 = wb3.GetSheetAt(0) as XSSFSheet;
            s2 = wb3.GetSheetAt(1) as XSSFSheet;
            s3 = wb3.GetSheetAt(2) as XSSFSheet;
            s4 = wb3.GetSheetAt(3) as XSSFSheet;
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

            wb3.Close();
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
            CellRangeAddress cra = CellRangeAddress.ValueOf("C2:D3");

            // No print Settings before repeating
            XSSFSheet s1 = wb1.CreateSheet() as XSSFSheet;
            Assert.AreEqual(false, s1.GetCTWorksheet().IsSetPageSetup());
            Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());
            s1.RepeatingColumns = (cra);
            s1.RepeatingRows = (cra);

            Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageSetup());
            Assert.AreEqual(true, s1.GetCTWorksheet().IsSetPageMargins());

            IPrintSetup ps1 = s1.PrintSetup;
            Assert.AreEqual(false, ps1.ValidSettings);
            Assert.AreEqual(false, ps1.Landscape);


            // Had valid print Settings before repeating
            XSSFSheet s2 = wb2.CreateSheet() as XSSFSheet;
            IPrintSetup ps2 = s2.PrintSetup;
            Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageSetup());
            Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageMargins());

            ps2.Landscape = (false);
            Assert.AreEqual(true, ps2.ValidSettings);
            Assert.AreEqual(false, ps2.Landscape);

            s2.RepeatingColumns = (cra);
            s2.RepeatingRows = (cra);

            ps2 = s2.PrintSetup;
            Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageSetup());
            Assert.AreEqual(true, s2.GetCTWorksheet().IsSetPageMargins());
            Assert.AreEqual(true, ps2.ValidSettings);
            Assert.AreEqual(false, ps2.Landscape);

            wb1.Close();
            wb2.Close();
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
            blueStyle.FillForegroundColor = (IndexedColors.Aqua.Index);
            blueStyle.FillPattern = (FillPattern.SolidForeground);
            Assert.AreEqual(1, blueStyle.Index);

            ICellStyle pinkStyle = wb.CreateCellStyle();
            pinkStyle.FillForegroundColor = (IndexedColors.Pink.Index);
            pinkStyle.FillPattern = (FillPattern.SolidForeground);
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

            wb.Close();
        }

        /**
         * Repeatedly writing a file.
         * Something with the SharedStringsTable currently breaks...
         */
        [Test]
        public void Test46662()
        {
            // New file
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            // Simple file
            XSSFWorkbook wb2 = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            XSSFTestDataSamples.WriteOutAndReadBack(wb2);
            XSSFTestDataSamples.WriteOutAndReadBack(wb2);
            XSSFTestDataSamples.WriteOutAndReadBack(wb2);
            wb2.Close();

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
            //       Assert.AreEqual("FFEEECE1", wb.Theme.GetThemeColor(2).ARGBHex);
            //       Assert.AreEqual("FF1F497D", wb.Theme.GetThemeColor(3).ARGBHex);

            // Finally check the colours on the styles
            // TODO fix
            //       Assert.AreEqual("FFEEECE1", cA4_EEECE1.CellStyle.GetFillForegroundXSSFColor().ARGBHex);
            //       Assert.AreEqual("FF1F497D", cA5_1F497D.CellStyle.GetFillForegroundXSSFColor().ARGBHex);

            wb.Close();
        }
        [Test]
        public void Test51470()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51470.xlsx");
            XSSFSheet sh0 = wb.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sh1 = wb.CloneSheet(0) as XSSFSheet;
            List<POIXMLDocumentPart.RelationPart> rels0 = sh0.RelationParts;
            List<POIXMLDocumentPart.RelationPart> rels1 = sh1.RelationParts;
            Assert.AreEqual(1, rels0.Count);
            Assert.AreEqual(1, rels1.Count);

            PackageRelationship pr0 = rels0[0].Relationship;
            PackageRelationship pr1 = rels1[0].Relationship;

            Assert.AreEqual(pr0.TargetMode, pr1.TargetMode);
            Assert.AreEqual(pr0.TargetUri, pr1.TargetUri);
            POIXMLDocumentPart doc0 = rels0[0].DocumentPart;
            POIXMLDocumentPart doc1 = rels1[0].DocumentPart;

            Assert.AreEqual(doc0, doc1);
            wb.Close();
        }

        /**
         * Add comments to Sheet 1, when Sheet 2 already has
         *  comments (so /xl/comments1.xml is taken)
         */
        [Test]
        public void Test51850()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("51850.xlsx");
            XSSFSheet sh1 = wb1.GetSheetAt(0) as XSSFSheet;
            XSSFSheet sh2 = wb1.GetSheetAt(1) as XSSFSheet;

            // Sheet 2 has comments
            Assert.IsNotNull(sh2.GetCommentsTable(false));
            Assert.AreEqual(1, sh2.GetCommentsTable(false).GetNumberOfComments());

            // Sheet 1 doesn't (yet)
            Assert.IsNull(sh1.GetCommentsTable(false));

            // Try to add comments to Sheet 1
            ICreationHelper factory = wb1.GetCreationHelper();
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

            anchor = factory.CreateClientAnchor();
            anchor.Col1 = (0);
            anchor.Col2 = (4);
            anchor.Row1 = (1);
            anchor.Row2 = (1);

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
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            wb1.Close();

            sh1 = wb2.GetSheetAt(0) as XSSFSheet;
            sh2 = wb2.GetSheetAt(1) as XSSFSheet;

            // Check the comments
            Assert.IsNotNull(sh2.GetCommentsTable(false));
            Assert.AreEqual(1, sh2.GetCommentsTable(false).GetNumberOfComments());

            Assert.IsNotNull(sh1.GetCommentsTable(false));
            Assert.AreEqual(2, sh1.GetCommentsTable(false).GetNumberOfComments());

            wb2.Close();
        }

        /**
         * Sheet names with a , in them
         */
        [Test]
        public void Test51963()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51963.xlsx");
            ISheet sheet = wb.GetSheetAt(0);
            Assert.AreEqual("Abc,1", sheet.SheetName);

            XSSFName name = wb.GetName("Intekon.ProdCodes") as XSSFName;
            Assert.AreEqual("'Abc,1'!$A$1:$A$2", name.RefersToFormula);

            AreaReference ref1 = new AreaReference(name.RefersToFormula);
            Assert.AreEqual(0, ref1.FirstCell.Row);
            Assert.AreEqual(0, ref1.FirstCell.Col);
            Assert.AreEqual(1, ref1.LastCell.Row);
            Assert.AreEqual(0, ref1.LastCell.Col);

            wb.Close();
        }

        /**
         * Sum across multiple workbooks
         *  eg =SUM($Sheet1.C1:$Sheet4.C1)
         * DISABLED As we can't currently Evaluate these
         */
         [Ignore("by poi")]
         [Test]
        public void Test48703()
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

            wb.Close();
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

            wb.Close();
        }

        /**
         * Bug 53101:
         */
        [Test]
        public void Test5301()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53101.xlsx");
            IFormulaEvaluator Evaluator =
                    wb.GetCreationHelper().CreateFormulaEvaluator();
            // A1: SUM(B1: IZ1)
            double a1Value =
                    Evaluator.Evaluate(wb.GetSheetAt(0).GetRow(0).GetCell(0)).NumberValue;

            // Assert
            Assert.AreEqual(259.0, a1Value, 0.0);

            // KY: SUM(B1: IZ1)
            double ky1Value =
                    Evaluator.Evaluate(wb.GetSheetAt(0).GetRow(0).GetCell(310)).NumberValue;

            // Assert
            Assert.AreEqual(259.0, a1Value, 0.0);

            wb.Close();
        }
        private class Function54436 : Function
        {
            public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
            {
                return ErrorEval.NA;
            }
        }
        [Test]
        public void Test54436()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("54436.xlsx");
            if (!WorkbookEvaluator.GetSupportedFunctionNames().Contains("GETPIVOTDATA"))
            {
                Function func = new Function54436();

                WorkbookEvaluator.RegisterFunction("GETPIVOTDATA", func);
            }
            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            wb.Close();
        }

        /**
         * Password Protected .xlsx files should give a helpful
         *  error message when called via WorkbookFactory.
         * (You need to supply a password explicitly for them)
         */
        [Test]
        public void Test55692()
        {
            Stream inpA = POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx");
            Stream inpB = POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx");
            Stream inpC = POIDataSamples.GetPOIFSInstance().OpenResourceAsStream("protect.xlsx");

            // Directly on a Stream
            try
            {
                WorkbookFactory.Create(inpA);
                Assert.Fail("Should've raised a EncryptedDocumentException error");
            }
            catch (EncryptedDocumentException e) { }

            // Via a POIFSFileSystem
            POIFSFileSystem fsP = new POIFSFileSystem(inpB);
            try
            {
                WorkbookFactory.Create(fsP);
                Assert.Fail("Should've raised a EncryptedDocumentException error");
            }
            catch (EncryptedDocumentException e) { }

            // Via a NPOIFSFileSystem
            NPOIFSFileSystem fsNP = new NPOIFSFileSystem(inpC);
            try
            {
                WorkbookFactory.Create(fsNP);
                Assert.Fail("Should've raised a EncryptedDocumentException error");
            }
            catch (EncryptedDocumentException e) { }
        }
        [Test]
        public void Bug53282()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53282b.xlsx");
            ICell c = wb.GetSheetAt(0).GetRow(1).GetCell(0);
            Assert.AreEqual("#@_#", c.StringCellValue);

            //with .net new Uri("mailto:#@_#") is valid, but java think it invalid, http://invalid.uri,
            //excel does nothing, it still show string "#@_#" 
            //should we add more validation to valid mail address in method PackagingUriHelper.ParseUri(string, UriKind)
            Assert.AreEqual("mailto:#@_#", c.Hyperlink.Address);

            wb.Close();
        }

        /**
         * Was giving NullPointerException
         * at NPOI.XSSF.UserModel.XSSFWorkbook.onDocumentRead
         * due to a lack of Styles Table
         */
        [Test]
        public void Bug56278()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56278.xlsx");
            Assert.AreEqual(0, wb.GetSheetIndex("Market Rates"));

            // Save and re-check
            IWorkbook nwb = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.AreEqual(0, nwb.GetSheetIndex("Market Rates"));

            wb.Close();
            nwb.Close();
        }

        [Test]
        public void Bug56315()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56315.xlsx");
            ICell c = wb.GetSheetAt(0).GetRow(1).GetCell(0);
            CellValue cv = wb.GetCreationHelper().CreateFormulaEvaluator().Evaluate(c);
            double rounded = cv.NumberValue;
            Assert.AreEqual(0.1, rounded, 0.0);

            wb.Close();
        }
        [Test]
        public void Bug56468()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFRow row = sheet.CreateRow(0) as XSSFRow;
            XSSFCell cell = row.CreateCell(0) as XSSFCell;
            cell.SetCellValue("Hi");
            sheet.RepeatingRows = (new CellRangeAddress(0, 0, 0, 0));

            MemoryStream bos = new MemoryStream(8096);
            wb.Write(bos);
            byte[] firstSave = bos.ToArray();
            //using (FileStream fs = new FileStream("d:\\save1.xlsx", FileMode.Create, FileAccess.ReadWrite))
            //{
            //    fs.Write(firstSave, 0, firstSave.Length);
            //    fs.Flush();
            //}

            MemoryStream bos2 = new MemoryStream(8096);
            wb.Write(bos2);
            byte[] secondSave = bos2.ToArray();
            //using (FileStream fs2 = new FileStream("d:\\save2.xlsx", FileMode.Create, FileAccess.ReadWrite))
            //{
            //    fs2.Write(secondSave, 0, secondSave.Length);
            //    fs2.Flush();
            //}


            Assert.That(firstSave, new EqualConstraint(secondSave),
                "Had: \n" + Arrays.ToString(firstSave) + " and \n" + Arrays.ToString(secondSave));

            wb.Close();
        }
        /**
         * ISO-8601 style cell formats with a T in them, eg
         * cell format of "yyyy-MM-ddTHH:mm:ss"
         */
        [Test]
        public void Bug54034()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("54034.xlsx");
            ISheet sheet = wb.GetSheet("Sheet1");
            IRow row = sheet.GetRow(1);
            ICell cell = row.GetCell(2);
            Assert.IsTrue(DateUtil.IsCellDateFormatted(cell));

            DataFormatter fmt = new DataFormatter();
            Assert.AreEqual("yyyy\\-mm\\-dd\\Thh:mm", cell.CellStyle.GetDataFormatString());
            Assert.AreEqual("2012-08-08T22:59", fmt.FormatCellValue(cell));

            wb.Close();
        }

        [Test]
        public void TestBug53798XLSX()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53798_ShiftNegative_TMPL.xlsx");
            FileInfo xlsOutput = TempFile.CreateTempFile("testBug53798", ".xlsx");
            bug53798Work(wb, xlsOutput);

            wb.Close();
        }


        [Ignore("Shifting rows is not yet implemented in SXSSFSheet")]
        [Test]
        public void TestBug53798XLSXStream()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53798_shiftNegative_TMPL.xlsx");
            FileInfo xlsOutput = TempFile.CreateTempFile("testBug53798", ".xlsx");
            SXSSFWorkbook wb2 = new SXSSFWorkbook(wb);
            bug53798Work(wb2, xlsOutput);
            wb2.Close();
            wb.Close();
        }

        [Test]
        public void TestBug53798XLS()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("53798_ShiftNegative_TMPL.xls");
            FileInfo xlsOutput = TempFile.CreateTempFile("testBug53798", ".xls");
            bug53798Work(wb, xlsOutput);

            wb.Close();
        }
        /**
         * SUMIF was throwing a NPE on some formulas
         */
        [Test]
        public void TestBug56420SumIfNPE()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56420.xlsx");

            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet sheet = wb.GetSheetAt(0);
            IRow r = sheet.GetRow(2);
            ICell c = r.GetCell(2);
            Assert.AreEqual("SUMIF($A$1:$A$4,A3,$B$1:$B$4)", c.CellFormula);
            evaluator.EvaluateInCell(c);

            ICell eval = evaluator.EvaluateInCell(c);
            Assert.AreEqual(0.0, eval.NumericCellValue, 0.0001);

            wb.Close();
        }
        private void bug53798Work(IWorkbook wb, FileInfo xlsOutput)
        {
            ISheet testSheet = wb.GetSheetAt(0);

            testSheet.ShiftRows(2, 2, 1);

            saveAndReloadReport(wb, xlsOutput);

            // 1) corrupted xlsx (unreadable data in the first row of a Shifted group) already comes about
            // when Shifted by less than -1 negative amount (try -2)
            testSheet.ShiftRows(3, 3, -1);

            saveAndReloadReport(wb, xlsOutput);

            testSheet.ShiftRows(2, 2, 1);

            saveAndReloadReport(wb, xlsOutput);

            // 2) attempt to create a new row IN PLACE of a Removed row by a negative shift causes corrupted
            // xlsx file with  unreadable data in the negative Shifted row.
            // NOTE it's ok to create any other row.
            IRow newRow = testSheet.CreateRow(3);

            saveAndReloadReport(wb, xlsOutput);

            ICell newCell = newRow.CreateCell(0);

            saveAndReloadReport(wb, xlsOutput);

            newCell.SetCellValue("new Cell in row " + newRow.RowNum);

            saveAndReloadReport(wb, xlsOutput);

            // 3) once a negative shift has been made any attempt to shift another group of rows
            // (note: outside of previously negative Shifted rows) by a POSITIVE amount causes POI exception:
            // org.apache.xmlbeans.impl.values.XmlValueDisconnectedException.
            // NOTE: another negative shift on another group of rows is successful, provided no new rows in
            // place of previously Shifted rows were attempted to be Created as explained above.
            testSheet.ShiftRows(6, 7, 1);   // -- CHANGE the shift to positive once the behaviour of
            // the above has been tested

            saveAndReloadReport(wb, xlsOutput);
        }

        /**
         * XSSFCell.typeMismatch on certain blank cells when formatting
         *  with DataFormatter
         */
        [Test]
        public void Bug56702()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56702.xlsx");

            ISheet sheet = wb.GetSheetAt(0);

            // Get wrong cell by row 8 & column 7
            ICell cell = sheet.GetRow(8).GetCell(7);
            Assert.AreEqual(CellType.Numeric, cell.CellType);

            // Check the value - will be zero as it is <c><v/></c>
            Assert.AreEqual(0.0, cell.NumericCellValue, 0.001);

            // Try to format
            DataFormatter formatter = new DataFormatter();
            formatter.FormatCellValue(cell);

            // Check the formatting
            Assert.AreEqual("0", formatter.FormatCellValue(cell));

            wb.Close();
        }

        /**
         * Formulas which reference named ranges, either in other
         *  sheets, or workbook scoped but in other workbooks.
         * Currently failing with errors like
         * NPOI.SS.Formula.FormulaParseException: Cell reference expected After sheet name at index 9
         * NPOI.SS.Formula.FormulaParseException: Parse error near char 0 '[' in specified formula '[0]!NR_Global_B2'. Expected number, string, or defined name 
         */
        [Test]
        public void Bug56737()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56737.xlsx");

            // Check the named range defInitions
            IName nSheetScope = wb.GetName("NR_To_A1");
            IName nWBScope = wb.GetName("NR_Global_B2");

            Assert.IsNotNull(nSheetScope);
            Assert.IsNotNull(nWBScope);

            Assert.AreEqual("Defines!$A$1", nSheetScope.RefersToFormula);
            Assert.AreEqual("Defines!$B$2", nWBScope.RefersToFormula);

            // Check the different kinds of formulas
            ISheet s = wb.GetSheetAt(0);
            ICell cRefSName = s.GetRow(1).GetCell(3);
            ICell cRefWName = s.GetRow(2).GetCell(3);

            Assert.AreEqual("Defines!NR_To_A1", cRefSName.CellFormula);
            Assert.AreEqual("[0]!NR_Global_B2", cRefWName.CellFormula);

            // Try to Evaluate them
            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual("Test A1", eval.Evaluate(cRefSName).StringValue);
            Assert.AreEqual(142, (int)eval.Evaluate(cRefWName).NumberValue);

            // Try to Evaluate everything
            eval.EvaluateAll();

            wb.Close();
        }

        private void saveAndReloadReport(IWorkbook wb, FileInfo outFile)
        {
            // run some method on the font to verify if it is "disconnected" already
            //for(short i = 0;i < 256;i++)
            {
                IFont font = wb.GetFontAt((short)0);
                if (font is XSSFFont)
                {
                    XSSFFont xfont = (XSSFFont)wb.GetFontAt((short)0);
                    CT_Font ctFont = (CT_Font)xfont.GetCTFont();
                    Assert.AreEqual(0, ctFont.SizeOfBArray());
                }
            }

            FileStream fileOutStream = new FileStream(outFile.FullName, FileMode.Open, FileAccess.ReadWrite);
            wb.Write(fileOutStream);
            fileOutStream.Close();
            //System.out.Println("File \""+outFile.Name+"\" has been saved successfully");

            FileStream is1 = new FileStream(outFile.FullName, FileMode.Open, FileAccess.ReadWrite);
            try
            {
                IWorkbook newWB = null;
                try
                {
                    if (wb is XSSFWorkbook)
                    {
                        newWB = new XSSFWorkbook(is1);
                    }
                    else if (wb is HSSFWorkbook)
                    {
                        newWB = new HSSFWorkbook(is1);
                        //} else if(wb is SXSSFWorkbook) {
                        //    newWB = new SXSSFWorkbook(new XSSFWorkbook(is1));
                    }
                    else
                    {
                        throw new InvalidOperationException("Unknown workbook: " + wb);
                    }
                    Assert.IsNotNull(newWB.GetSheet("test"));
                }
                finally
                {
                    if (newWB != null)
                    {
                        //newWB.Close();
                    }
                }
            }
            finally
            {
                is1.Close();
            }
        }

        [Test]
        public void TestBug56688_1()
        {
            XSSFWorkbook excel = XSSFTestDataSamples.OpenSampleWorkbook("56688_1.xlsx");
            //CheckValue(excel, "-1.0");  /* Not 0.0 because POI sees date "0" minus one month as invalid date, which is -1! */
            CheckValue(excel, "-1");

            excel.Close();
        }

        [Test]
        public void TestBug56688_2()
        {
            XSSFWorkbook excel = XSSFTestDataSamples.OpenSampleWorkbook("56688_2.xlsx");
            CheckValue(excel, "#VALUE!");
            excel.Close();
        }

        [Test]
        public void TestBug56688_3()
        {
            XSSFWorkbook excel = XSSFTestDataSamples.OpenSampleWorkbook("56688_3.xlsx");
            CheckValue(excel, "#VALUE!");
            excel.Close();
        }

        [Test]
        public void TestBug56688_4()
        {
            XSSFWorkbook excel = XSSFTestDataSamples.OpenSampleWorkbook("56688_4.xlsx");

            Calendar calendar = new GregorianCalendar(GregorianCalendarTypes.USEnglish);
            DateTime time = calendar.AddMonths(DateTime.Now, 2);
            double excelDate = DateUtil.GetExcelDate(time);
            NumberEval eval = new NumberEval(Math.Floor(excelDate));
            //CheckValue(excel, eval.StringValue + ".0");
            CheckValue(excel, eval.StringValue);
            excel.Close();
        }

        /**
         * New hyperlink with no initial cell reference, still need
         *  to be able to change it
         */
        [Test]
        public void TestBug56527()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFCreationHelper creationHelper = wb.GetCreationHelper() as XSSFCreationHelper;
            XSSFHyperlink hyperlink;

            // Try with a cell reference
            hyperlink = creationHelper.CreateHyperlink(HyperlinkType.Url) as XSSFHyperlink;
            sheet.AddHyperlink(hyperlink);
            hyperlink.Address = (/*setter*/"http://myurl");
            hyperlink.SetCellReference(/*setter*/"B4");
            Assert.AreEqual(3, hyperlink.FirstRow);
            Assert.AreEqual(1, hyperlink.FirstColumn);
            Assert.AreEqual(3, hyperlink.LastRow);
            Assert.AreEqual(1, hyperlink.LastColumn);

            // Try with explicit rows / columns
            hyperlink = creationHelper.CreateHyperlink(HyperlinkType.Url) as XSSFHyperlink;
            sheet.AddHyperlink(hyperlink);
            hyperlink.Address = (/*setter*/"http://myurl");
            hyperlink.FirstRow = (/*setter*/5);
            hyperlink.FirstColumn = (/*setter*/3);

            Assert.AreEqual(5, hyperlink.FirstRow);
            Assert.AreEqual(3, hyperlink.FirstColumn);
            Assert.AreEqual(5, hyperlink.LastRow);
            Assert.AreEqual(3, hyperlink.LastColumn);

            wb.Close();
        }

        /**
         * Shifting rows with a formula that references a 
         * function in another file
         */
        [Test]
        public void Bug56502()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56502.xlsx");
            ISheet sheet = wb.GetSheetAt(0);

            ICell cFunc = sheet.GetRow(3).GetCell(0);
            Assert.AreEqual("[1]!LUCANET(\"Ist\")", cFunc.CellFormula);
            ICell cRef = sheet.GetRow(3).CreateCell(1);
            cRef.CellFormula = (/*setter*/"A3");

            // Shift it down one row
            sheet.ShiftRows(1, sheet.LastRowNum, 1);

            // Check the new formulas: Function won't Change, Reference will
            cFunc = sheet.GetRow(4).GetCell(0);
            Assert.AreEqual("[1]!LUCANET(\"Ist\")", cFunc.CellFormula);
            cRef = sheet.GetRow(4).GetCell(1);
            Assert.AreEqual("A4", cRef.CellFormula);

            wb.Close();
        }

        [Test]
        public void Bug54764()
        {
            OPCPackage pkg = XSSFTestDataSamples.OpenSamplePackage("54764.xlsx");

            // Check the core properties - will be found but empty, due
            //  to the expansion being too much to be considered valid
            POIXMLProperties props = new POIXMLProperties(pkg);
            Assert.AreEqual(null, props.CoreProperties.Title);
            Assert.AreEqual(null, props.CoreProperties.Subject);
            Assert.AreEqual(null, props.CoreProperties.Description);

            // Now check the spreadsheet itself

            try
            {
                new XSSFWorkbook(pkg).Close();
                Assert.Fail("Should fail as too much expansion occurs");
            }
            catch (POIXMLException)
            {
                //Expected
            }
            pkg.Close();

            // Try with one with the entities in the Content Types
            try
            {
                XSSFTestDataSamples.OpenSamplePackage("54764-2.xlsx").Close();
                Assert.Fail("Should fail as too much expansion occurs");
            }
            catch (Exception)
            {
                // Expected
            }

            // Check we can still parse valid files after all that
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            Assert.AreEqual(3, wb.NumberOfSheets);

            wb.Close();
        }

        /**
         * CTDefinedNamesImpl should be included in the smaller
         *  poi-ooxml-schemas jar
         */
        [Test]
        public void Bug57176()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57176.xlsx");
            CT_DefinedNames definedNames = wb.GetCTWorkbook().definedNames;
            List<CT_DefinedName> definedNameList = definedNames.definedName;
            foreach (CT_DefinedName defName in definedNameList)
            {
                Assert.IsNotNull(defName.name);
                Assert.IsNotNull(defName.Value);
            }
            Assert.AreEqual("TestDefinedName", definedNameList[(0)].name);

            wb.Close();
        }

        /**
         * .xlsb files are not supported, but we should generate a helpful
         *  error message if given one
         */
        [Test]
        public void Bug56800_xlsb()
        {
            // Can be opened at the OPC level
            OPCPackage pkg = XSSFTestDataSamples.OpenSamplePackage("Simple.xlsb");

            // XSSF Workbook gives helpful error
            try
            {
                new XSSFWorkbook(pkg).Close();
                Assert.Fail(".xlsb files not supported");
            }
            catch (XLSBUnsupportedException e)
            {
                // Good, detected and warned
            }

            // Workbook Factory gives helpful error on package
            try
            {
                WorkbookFactory.Create(pkg).Close();
                Assert.Fail(".xlsb files not supported");
            }
            catch (XLSBUnsupportedException e)
            {
                // Good, detected and warned
            }

            // Workbook Factory gives helpful error on file
            FileInfo xlsbFile = HSSFTestDataSamples.GetSampleFile("Simple.xlsb");
            try
            {
                WorkbookFactory.Create(xlsbFile.FullName).Close();
                Assert.Fail(".xlsb files not supported");
            }
            catch (XLSBUnsupportedException e)
            {
                // Good, detected and warned
            }

            pkg.Close();
        }

        private void CheckValue(XSSFWorkbook excel, String expect)
        {
            XSSFFormulaEvaluator Evaluator = new XSSFFormulaEvaluator(excel);
            Evaluator.EvaluateAll();

            XSSFCell cell = excel.GetSheetAt(0).GetRow(1).GetCell(1) as XSSFCell;
            CellValue value = Evaluator.Evaluate(cell);

            Assert.AreEqual(expect, value.FormatAsString());
        }

        [Test]
        public void TestBug57196()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57196.xlsx");
            ISheet sheet = wb.GetSheet("Feuil1");
            IRow mod = sheet.GetRow(1);
            mod.GetCell(1).SetCellValue(3);
            HSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);
            //        FileOutputStream fileOutput = new FileOutputStream("/tmp/57196.xlsx");
            //        wb.Write(fileOutput);
            //        fileOutput.Close();
            wb.Close();
        }

        [Test]
        public void Test57196_Detail()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet("Sheet1") as XSSFSheet;
            XSSFRow row = sheet.CreateRow(0) as XSSFRow;
            XSSFCell cell = row.CreateCell(0) as XSSFCell;
            cell.CellFormula = (/*setter*/"DEC2HEX(HEX2DEC(O8)-O2+D2)");
            XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);

            Assert.IsNotNull(cv);
            wb.Close();
        }

        [Test]
        public void Test57196_Detail2()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet("Sheet1") as XSSFSheet;
            XSSFRow row = sheet.CreateRow(0) as XSSFRow;
            XSSFCell cell = row.CreateCell(0) as XSSFCell;
            cell.CellFormula = (/*setter*/"DEC2HEX(O2+D2)");
            XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(wb);
            CellValue cv = fe.Evaluate(cell);

            Assert.IsNotNull(cv);
            wb.Close();
        }

        [Test]
        public void Test57196_WorkbookEvaluator()
        {
            //Environment.SetEnvironmentVariable("NPOI.UTIL.POILogger", "NPOI.UTIL.SystemOutLogger");
            //Environment.SetEnvironmentVariable("poi.log.level", "3");
            try
            {
                XSSFWorkbook wb = new XSSFWorkbook();
                XSSFSheet sheet = wb.CreateSheet("Sheet1") as XSSFSheet;
                XSSFRow row = sheet.CreateRow(0) as XSSFRow;
                XSSFCell cell = row.CreateCell(0) as XSSFCell;

                cell.SetCellValue("0");
                cell = row.CreateCell(1) as XSSFCell;
                cell.SetCellValue(0);
                cell = row.CreateCell(2) as XSSFCell;
                cell.SetCellValue(0);


                // simple formula worked
                cell.CellFormula = (/*setter*/"DEC2HEX(O2+D2)");

                WorkbookEvaluator workbookEvaluator = new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(wb), null, null);
                workbookEvaluator.DebugEvaluationOutputForNextEval = (/*setter*/true);
                workbookEvaluator.Evaluate(new XSSFEvaluationCell(cell));

                // this already failed! Hex2Dec did not correctly handle RefEval
                cell.CellFormula = (/*setter*/"HEX2DEC(O8)");
                workbookEvaluator.ClearAllCachedResultValues();

                workbookEvaluator = new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(wb), null, null);
                workbookEvaluator.DebugEvaluationOutputForNextEval = (/*setter*/true);
                workbookEvaluator.Evaluate(new XSSFEvaluationCell(cell));

                // slightly more complex one failed
                cell.CellFormula = (/*setter*/"HEX2DEC(O8)-O2+D2");
                workbookEvaluator.ClearAllCachedResultValues();

                workbookEvaluator = new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(wb), null, null);
                workbookEvaluator.DebugEvaluationOutputForNextEval = (/*setter*/true);
                workbookEvaluator.Evaluate(new XSSFEvaluationCell(cell));

                // more complicated failed
                cell.CellFormula = (/*setter*/"DEC2HEX(HEX2DEC(O8)-O2+D2)");
                workbookEvaluator.ClearAllCachedResultValues();

                workbookEvaluator.DebugEvaluationOutputForNextEval = (/*setter*/true);
                workbookEvaluator.Evaluate(new XSSFEvaluationCell(cell));

                // what other similar functions
                cell.CellFormula = (/*setter*/"DEC2BIN(O8)-O2+D2");
                workbookEvaluator.ClearAllCachedResultValues();

                workbookEvaluator = new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(wb), null, null);
                workbookEvaluator.DebugEvaluationOutputForNextEval = (/*setter*/true);
                workbookEvaluator.Evaluate(new XSSFEvaluationCell(cell));

                // what other similar functions
                cell.CellFormula = (/*setter*/"DEC2BIN(A1)");
                workbookEvaluator.ClearAllCachedResultValues();

                workbookEvaluator = new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(wb), null, null);
                workbookEvaluator.DebugEvaluationOutputForNextEval = (/*setter*/true);
                workbookEvaluator.Evaluate(new XSSFEvaluationCell(cell));

                // what other similar functions
                cell.CellFormula = (/*setter*/"BIN2DEC(B1)");
                workbookEvaluator.ClearAllCachedResultValues();

                workbookEvaluator = new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(wb), null, null);
                workbookEvaluator.DebugEvaluationOutputForNextEval = (/*setter*/true);
                workbookEvaluator.Evaluate(new XSSFEvaluationCell(cell));
                wb.Close();
            }
            finally
            {
                //System.ClearProperty("NPOI.UTIL.POILogger");
                //System.ClearProperty("poi.log.level");
            }
        }


        /**
         * A .xlsx file with no Shared Strings table should open fine
         *  in Read-only mode
         */
        [Test]
        public void Bug57482()
        {
            foreach (PackageAccess access in new PackageAccess[] {
                PackageAccess.READ_WRITE, PackageAccess.READ
        })
            {
                FileInfo file = HSSFTestDataSamples.GetSampleFile("57482-OnlyNumeric.xlsx");
                OPCPackage pkg = OPCPackage.Open(file, access);
                try
                {
                    XSSFWorkbook wb = new XSSFWorkbook(pkg);
                    Assert.IsNotNull(wb.GetSharedStringSource());
                    Assert.AreEqual(0, wb.GetSharedStringSource().Count);

                    DataFormatter fmt = new DataFormatter();
                    XSSFSheet s = wb.GetSheetAt(0) as XSSFSheet;
                    Assert.AreEqual("1", fmt.FormatCellValue(s.GetRow(0).GetCell(0)));
                    Assert.AreEqual("11", fmt.FormatCellValue(s.GetRow(0).GetCell(1)));
                    Assert.AreEqual("5", fmt.FormatCellValue(s.GetRow(4).GetCell(0)));

                    // Add a text cell
                    s.GetRow(0).CreateCell(3).SetCellValue("Testing");
                    Assert.AreEqual("Testing", fmt.FormatCellValue(s.GetRow(0).GetCell(3)));

                    // Try to Write-out and read again, should only work
                    //  in Read-write mode, not Read-only mode
                    try
                    {
                        wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
                        if (access == PackageAccess.READ)
                            Assert.Fail("Shouln't be able to write from Read-only mode");
                    }
                    catch (InvalidOperationException e)
                    {
                        if (access == PackageAccess.READ)
                        {
                            // Expected
                        }
                        else
                        {
                            // Shouldn't occur in Write-mode
                            throw e;
                        }
                    }

                    // Check again
                    s = wb.GetSheetAt(0) as XSSFSheet;
                    Assert.AreEqual("1", fmt.FormatCellValue(s.GetRow(0).GetCell(0)));
                    Assert.AreEqual("11", fmt.FormatCellValue(s.GetRow(0).GetCell(1)));
                    Assert.AreEqual("5", fmt.FormatCellValue(s.GetRow(4).GetCell(0)));
                    Assert.AreEqual("Testing", fmt.FormatCellValue(s.GetRow(0).GetCell(3)));

                }
                finally
                {
                    pkg.Revert();
                }
            }
        }

        /**
     * "Unknown error type: -60" fetching formula error value
     */
        [Test]
        public void Bug57535()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57535.xlsx");
            IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            Evaluator.ClearAllCachedResultValues();

            ISheet sheet = wb.GetSheet("Sheet1");
            ICell cell = sheet.GetRow(5).GetCell(4);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual("E4+E5", cell.CellFormula);

            CellValue value = Evaluator.Evaluate(cell);
            Assert.AreEqual(CellType.Error, value.CellType);
            Assert.AreEqual(-60, value.ErrorValue);

            Assert.AreEqual("~CIRCULAR~REF~", FormulaError.ForInt(value.ErrorValue).String);
            Assert.AreEqual("CIRCULAR_REF", FormulaError.ForInt(value.ErrorValue).ToString());
            wb.Close();
        }

        [Test]
        public void Test57165()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");
            try
            {
                RemoveAllSheetsBut(3, wb);
                wb.CloneSheet(0); // Throws exception here
                wb.SetSheetName(1, "New Sheet");
                //saveWorkbook(wb, fileName);

                XSSFWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
                wbBack.Close();
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void Test57165_Create()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57171_57163_57165.xlsx");
            try
            {
                RemoveAllSheetsBut(3, wb);
                wb.CreateSheet("newsheet"); // Throws exception here
                wb.SetSheetName(1, "New Sheet");
                //saveWorkbook(wb, fileName);

                XSSFWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
                wbBack.Close();
            }
            finally
            {
                wb.Close();
            }
        }


        private static void RemoveAllSheetsBut(int sheetIndex, IWorkbook wb)
        {
            int sheetNb = wb.NumberOfSheets;
            // Move this sheet at the first position
            wb.SetSheetOrder(wb.GetSheetName(sheetIndex), 0);
            for (int sn = sheetNb - 1; sn > 0; sn--)
            {
                wb.RemoveSheetAt(sn);
            }
        }

        /**
     * Sums 2 plus the cell at the left, indirectly to avoid reference
     * problems when deleting columns, conditionally to stop recursion
     */
        private static String FORMULA1 =
                "IF( INDIRECT( ADDRESS( ROW(), COLUMN()-1 ) ) = 0, 0,"
                        + "INDIRECT( ADDRESS( ROW(), COLUMN()-1 ) ) ) + 2";

        /**
         * Sums 2 plus the upper cell, indirectly to avoid reference
         * problems when deleting rows, conditionally to stop recursion
         */
        private static String FORMULA2 =
                "IF( INDIRECT( ADDRESS( ROW()-1, COLUMN() ) ) = 0, 0,"
                        + "INDIRECT( ADDRESS( ROW()-1, COLUMN() ) ) ) + 2";

        /**
         * Expected:

         * [  0][  2][  4]
         * @
         */
        [Test]
        public void TestBug56820_Formula1()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
                ISheet sh = wb.CreateSheet();

                sh.CreateRow(0).CreateCell(0).SetCellValue(0.0d);
                ICell formulaCell1 = sh.GetRow(0).CreateCell(1);
                ICell formulaCell2 = sh.GetRow(0).CreateCell(2);
                formulaCell1.CellFormula = (/*setter*/FORMULA1);
                formulaCell2.CellFormula = (/*setter*/FORMULA1);

                double A1 = Evaluator.Evaluate(formulaCell1).NumberValue;
                double A2 = Evaluator.Evaluate(formulaCell2).NumberValue;

                Assert.AreEqual(2, A1, 0);
                Assert.AreEqual(4, A2, 0);  //<-- FAILS EXPECTATIONS
            }
            finally
            {
                wb.Close();
            }
        }

        /**
         * Expected:

         * [  0] <- number
         * [  2] <- formula
         * [  4] <- formula
         * @
         */
        [Test]
        public void TestBug56820_Formula2()
        {
            IWorkbook wb = new XSSFWorkbook();
            try
            {
                IFormulaEvaluator Evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
                ISheet sh = wb.CreateSheet();

                sh.CreateRow(0).CreateCell(0).SetCellValue(0.0d);
                ICell formulaCell1 = sh.CreateRow(1).CreateCell(0);
                ICell formulaCell2 = sh.CreateRow(2).CreateCell(0);
                formulaCell1.CellFormula = (/*setter*/FORMULA2);
                formulaCell2.CellFormula = (/*setter*/FORMULA2);

                double A1 = Evaluator.Evaluate(formulaCell1).NumberValue;
                double A2 = Evaluator.Evaluate(formulaCell2).NumberValue; //<-- FAILS EVALUATION

                Assert.AreEqual(2, A1, 0);
                Assert.AreEqual(4, A2, 0);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void Test56467()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("picture.xlsx");
            try
            {
                ISheet orig = wb.GetSheetAt(0);
                Assert.IsNotNull(orig);

                ISheet sheet = wb.CloneSheet(0);
                IDrawing Drawing = sheet.CreateDrawingPatriarch();
                foreach (XSSFShape shape in ((XSSFDrawing)Drawing).GetShapes())
                {
                    if (shape is XSSFPicture)
                    {
                        XSSFPictureData pictureData = ((XSSFPicture)shape).PictureData as XSSFPictureData;
                        Assert.IsNotNull(pictureData);
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }
        /**
     * OOXML-Strict files
     * Not currently working - namespace mis-match from XMLBeans
     */
        [Ignore("XMLBeans namespace mis-match on ooxml-strict files")]
        [Test]
        public void Test57699()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("sample.strict.xlsx");
            Assert.AreEqual(3, wb.NumberOfSheets);
            // TODO Check sheet contents
            // TODO Check formula Evaluation

            XSSFWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            Assert.AreEqual(3, wbBack.NumberOfSheets);
            // TODO Re-check sheet contents
            // TODO Re-check formula Evaluation

            wb.Close();
            wbBack.Close();
        }

        [Test]
        public void TestBug56295_MergeXlslsWithStyles()
        {
            XSSFWorkbook xlsToAppendWorkbook = XSSFTestDataSamples.OpenSampleWorkbook("56295.xlsx");
            XSSFSheet sheet = xlsToAppendWorkbook.GetSheetAt(0) as XSSFSheet;
            XSSFRow srcRow = sheet.GetRow(0) as XSSFRow;
            XSSFCell oldCell = srcRow.GetCell(0) as XSSFCell;
            XSSFCellStyle cellStyle = oldCell.CellStyle as XSSFCellStyle;

            CheckStyle(cellStyle);

            //        StylesTable table = xlsToAppendWorkbook.StylesSource;
            //        List<XSSFCellFill> Fills = table.Fills;
            //        System.out.Println("Having " + Fills.Count + " Fills");
            //        for(XSSFCellFill fill : Fills) {
            //        	System.out.Println("Fill: " + Fill.FillBackgroundColor + "/" + Fill.FillForegroundColor);
            //        }        

            xlsToAppendWorkbook.Close();

            XSSFWorkbook targetWorkbook = new XSSFWorkbook();
            XSSFSheet newSheet = targetWorkbook.CreateSheet(sheet.SheetName) as XSSFSheet;
            XSSFRow destRow = newSheet.CreateRow(0) as XSSFRow;
            XSSFCell newCell = destRow.CreateCell(0) as XSSFCell;

            //newCell.CellStyle.CloneStyleFrom(cellStyle);
            ICellStyle newCellStyle = targetWorkbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(cellStyle);
            newCell.CellStyle = (/*setter*/newCellStyle);
            CheckStyle(newCell.CellStyle as XSSFCellStyle);
            newCell.SetCellValue(oldCell.StringCellValue);

            //        OutputStream os = new FileOutputStream("output.xlsm");
            //        try {
            //        	targetWorkbook.Write(os);
            //        } finally {
            //        	os.Close();
            //        }

            XSSFWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(targetWorkbook) as XSSFWorkbook;
            XSSFCellStyle styleBack = wbBack.GetSheetAt(0).GetRow(0).GetCell(0).CellStyle as XSSFCellStyle;
            CheckStyle(styleBack);

            targetWorkbook.Close();
            wbBack.Close();
        }

        /**
         * Paragraph with property BuFont but none of the properties
         *  BuNone, BuChar, and BuAutoNum, used to trigger a NPE
         * Excel treats this as not-bulleted, so now do we
         */
        [Test]
        public void TestBug57826()  {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("57826.xlsx");

            Assert.IsTrue(workbook.NumberOfSheets >= 1, "no sheets in workbook");
            XSSFSheet sheet = workbook.GetSheetAt(0) as XSSFSheet;

            XSSFDrawing drawing = sheet.GetDrawingPatriarch();
            Assert.IsNotNull(drawing);

            List<XSSFShape> shapes = drawing.GetShapes();
            Assert.AreEqual(1, shapes.Count);
            Assert.IsTrue(shapes[0] is XSSFSimpleShape);

            XSSFSimpleShape shape = (XSSFSimpleShape)shapes[0];

            // Used to throw a NPE
            String text = shape.Text;

            // No bulleting info included
            Assert.AreEqual("test ok", text);
        
            workbook.Close();
        }

        private void CheckStyle(XSSFCellStyle cellStyle)
        {
            Assert.IsNotNull(cellStyle);
            Assert.AreEqual(0, cellStyle.FillForegroundColor);
            Assert.IsNotNull(cellStyle.FillForegroundXSSFColor);
            XSSFColor fgColor = cellStyle.FillForegroundColorColor as XSSFColor;
            Assert.IsNotNull(fgColor);
            Assert.AreEqual("FF00FFFF", fgColor.ARGBHex);

            Assert.AreEqual(0, cellStyle.FillBackgroundColor);
            Assert.IsNotNull(cellStyle.FillBackgroundXSSFColor);
            XSSFColor bgColor = cellStyle.FillBackgroundColorColor as XSSFColor;
            Assert.IsNotNull(bgColor);
            Assert.AreEqual("FF00FFFF", fgColor.ARGBHex);
        }

        [Test]
        public void Bug57642()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet s = wb.CreateSheet("TestSheet") as XSSFSheet;
            XSSFCell c = s.CreateRow(0).CreateCell(0) as XSSFCell;
            c.CellFormula = (/*setter*/"ISERROR(TestSheet!A1)");
            c = s.CreateRow(1).CreateCell(1) as XSSFCell;
            c.CellFormula = (/*setter*/"ISERROR(B2)");

            wb.SetSheetName(0, "CSN");
            c = s.GetRow(0).GetCell(0) as XSSFCell;
            Assert.AreEqual("ISERROR(CSN!A1)", c.CellFormula);
            c = s.GetRow(1).GetCell(1) as XSSFCell;
            Assert.AreEqual("ISERROR(B2)", c.CellFormula);
        }

        /**
         * .xlsx supports 64000 cell styles, the style indexes After
         *  32,767 must not be -32,768, then -32,767, -32,766
         *  long time test, run over 1 minute.
         */
        [Test, RunSerialyAndSweepTmpFiles]
        public void Bug57880()
        {
            Console.WriteLine("long time test, run over 1 minute.");
            int numStyles = 33000;
            XSSFWorkbook wb = new XSSFWorkbook();
            for (int i = 1; i < numStyles; i++)
            {
                // Create a style and use it
                XSSFCellStyle style = wb.CreateCellStyle() as XSSFCellStyle;
                Assert.AreEqual(i, style.UIndex);
            }
            Assert.AreEqual(numStyles, wb.NumCellStyles);

            // avoid OOM in gump run
            FileInfo file = XSSFTestDataSamples.WriteOutAndClose(wb, "bug57880");
            wb = null;
            // Garbage collection may happen here

            //// avoid zip bomb detection
            //double ratio = ZipSecureFile.getMinInflateRatio();
            //ZipSecureFile.setMinInflateRatio(0.00005);
            //wb = XSSFTestDataSamples.ReadBackAndDelete(file);
            //ZipSecureFile.setMinInflateRatio(ratio);

            //Assume identical cell styles aren't consolidated
            //If XSSFWorkbooks ever implicitly optimize/consolidate cell styles (such as when the workbook is written to disk)
            //then this unit test should be updated
            Assert.AreEqual(numStyles, wb.NumCellStyles);
            for (int i = 1; i < numStyles; i++)
            {
                XSSFCellStyle style = wb.GetCellStyleAt((short)i) as XSSFCellStyle;
                Assert.IsNotNull(style);
                Assert.AreEqual(i, style.UIndex);
            }

            wb.Close();

            Assert.AreEqual(0, Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.tmp").Length, "At Last: There are no temporary files.");
        }
        [Test]
        public void Test56574()
        {
            RunTest56574(false);
            RunTest56574(true);
        }

        private void RunTest56574(bool CreateRow)
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("56574.xlsx");

            ISheet sheet = wb.GetSheet("Func");
            Assert.IsNotNull(sheet);

            Dictionary<String, Object[]> data;
            data = new Dictionary<String, Object[]>();
            data.Add("1", new Object[] { "ID", "NAME", "LASTNAME" });
            data.Add("2", new Object[] { 2, "Amit", "Shukla" });
            data.Add("3", new Object[] { 1, "Lokesh", "Gupta" });
            data.Add("4", new Object[] { 4, "John", "Adwards" });
            data.Add("5", new Object[] { 2, "Brian", "Schultz" });

            var keyset = data.Keys;
            int rownum = 1;
            foreach (String key in keyset)
            {
                IRow row;
                if (CreateRow)
                {
                    row = sheet.CreateRow(rownum++);
                }
                else
                {
                    row = sheet.GetRow(rownum++);
                }
                Assert.IsNotNull(row);

                Object[] objArr = data[(key)];
                int cellnum = 0;
                foreach (Object obj in objArr)
                {
                    ICell cell = row.GetCell(cellnum);
                    if (cell == null)
                    {
                        cell = row.CreateCell(cellnum);
                    }
                    else
                    {
                        if (cell.CellType == CellType.Formula)
                        {
                            cell.CellFormula = (/*setter*/null);
                            cell.CellStyle.DataFormat = (/*setter*/(short)0);
                        }
                    }
                    if (obj is String)
                    {
                        cell.SetCellValue((String)obj);
                    }
                    else if (obj is int)
                    {
                        cell.SetCellValue((int)obj);
                    }
                    cellnum++;
                }
            }

            XSSFFormulaEvaluator.EvaluateAllFormulaCells((XSSFWorkbook)wb);
            wb.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            CalculationChain chain = ((XSSFWorkbook)wb).GetCalculationChain();
            foreach (CT_CalcCell calc in chain.GetCTCalcChain().c)
            {
                // A2 to A6 should be gone
                Assert.IsFalse(calc.r.Equals("A2"));
                Assert.IsFalse(calc.r.Equals("A3"));
                Assert.IsFalse(calc.r.Equals("A4"));
                Assert.IsFalse(calc.r.Equals("A5"));
                Assert.IsFalse(calc.r.Equals("A6"));
            }

            /*FileOutputStream out1 = new FileOutputStream(new File("C:\\temp\\56574.xlsx"));
            try {
                wb.Write(out1);
            } finally {
                out1.Close();
            }*/

            IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            ISheet sheetBack = wbBack.GetSheet("Func");
            Assert.IsNotNull(sheetBack);

            chain = ((XSSFWorkbook)wbBack).GetCalculationChain();
            foreach (CT_CalcCell calc in chain.GetCTCalcChain().c)
            {
                // A2 to A6 should be gone
                Assert.IsFalse(calc.r.Equals("A2"));
                Assert.IsFalse(calc.r.Equals("A3"));
                Assert.IsFalse(calc.r.Equals("A4"));
                Assert.IsFalse(calc.r.Equals("A5"));
                Assert.IsFalse(calc.r.Equals("A6"));
            }

            wbBack.Close();
            wb.Close();
        }

        /**
         * Excel 2007 generated Macro-Enabled .xlsm file
         */
        [Test]
        public void Bug57181()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57181.xlsm");
            Assert.AreEqual(9, wb.NumberOfSheets);
            wb.Close();
        }

        [Test]
        public void Bug52111()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("Intersection-52111-xssf.xlsx");
            ISheet s = wb.GetSheetAt(0);
            assertFormula(wb, s.GetRow(2).GetCell(0), "(C2:D3 D3:E4)", "4");
            assertFormula(wb, s.GetRow(6).GetCell(0), "Tabelle2!E:E Tabelle2!11:11", "5");
            assertFormula(wb, s.GetRow(8).GetCell(0), "Tabelle2!E:F Tabelle2!11:12", null);
            wb.Close();
        }

        private void assertFormula(IWorkbook wb, ICell intF, String expectedFormula, String expectedResultOrNull)
        {
            Assert.AreEqual(CellType.Formula, intF.CellType);
            if (null == expectedResultOrNull)
            {
                Assert.AreEqual(CellType.Error, intF.CachedFormulaResultType);
                expectedResultOrNull = "#VALUE!";
            }
            else
            {
                Assert.AreEqual(CellType.Numeric, intF.CachedFormulaResultType);
            }

            Assert.AreEqual(expectedFormula, intF.CellFormula);
            // Check we can evaluate it correctly
            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
            Assert.AreEqual(expectedResultOrNull, eval.Evaluate(intF).FormatAsString());
        }

        [Test]
        public void Test48962()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48962.xlsx");
            ISheet sh = wb.GetSheetAt(0);
            IRow row = sh.GetRow(1);
            ICell cell = row.GetCell(0);

            ICellStyle style = cell.CellStyle;
            Assert.IsNotNull(style);

            // color index
            Assert.AreEqual(64, style.FillBackgroundColor);
            XSSFColor color = ((XSSFCellStyle)style).FillBackgroundXSSFColor;
            Assert.IsNotNull(color);

            // indexed color
            Assert.AreEqual(64, color.Indexed);
            Assert.AreEqual(64, color.Index);

            // not an RGB color
            Assert.IsFalse(color.IsRGB);
            Assert.IsNull(color.RGB);

            wb.Close();
        }

        [Test]
        public void Test50755_workday_formula_example()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("50755_workday_formula_example.xlsx");
            ISheet sheet = wb.GetSheet("Sheet1");
            foreach (IRow aRow in sheet)
            {
                ICell cell = aRow.GetCell(1);
                if (cell.CellType == CellType.Formula)
                {
                    String formula = cell.CellFormula;
                    //System.out.println("formula: " + formula);
                    Assert.IsNotNull(formula);
                    Assert.IsTrue(formula.Contains("WORKDAY"));
                }
                else
                {
                    Assert.IsNotNull(cell.ToString());
                }
            }

            wb.Close();
        }

        [Test]
        public void Test51626()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51626.xlsx");
            Assert.IsNotNull(wb);
            wb.Close();

            Stream stream = HSSFTestDataSamples.OpenSampleFileStream("51626.xlsx");
            wb = WorkbookFactory.Create(stream);
            stream.Close();
            wb.Close();

            wb = XSSFTestDataSamples.OpenSampleWorkbook("51626_contact.xlsx");
            Assert.IsNotNull(wb);
            wb.Close();

            stream = HSSFTestDataSamples.OpenSampleFileStream("51626_contact.xlsx");
            wb = WorkbookFactory.Create(stream);
            stream.Close();
            wb.Close();

        }

        [Test]
        public void Test51451()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sh = wb.CreateSheet();

            IRow row = sh.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(239827342);

            ICellStyle style = wb.CreateCellStyle();
            //style.setHidden(false);
            IDataFormat excelFormat = wb.CreateDataFormat();
            style.DataFormat = (excelFormat.GetFormat("#,##0"));
            sh.SetDefaultColumnStyle(0, style);
            //        FileOutputStream out = new FileOutputStream("/tmp/51451.xlsx");
            //        wb.write(out);
            //        out.close();

            wb.Close();
        }

        [Test]
        public void Test53105()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53105.xlsx");
            Assert.IsNotNull(wb);

            // Act
            // evaluate SUM('Skye Lookup Input'!A4:XFD4), cells in range each contain "1"
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            double numericValue = evaluator.Evaluate(wb.GetSheetAt(0).GetRow(1).GetCell(0)).NumberValue;
            // Assert
            Assert.AreEqual(16384.0, numericValue, 0.0);

            wb.Close();
        }

        [Test]
        public void Test58315()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("58315.xlsx");
            ICell cell = wb.GetSheetAt(0).GetRow(0).GetCell(0);
            Assert.IsNotNull(cell);
            StringBuilder tmpCellContent = new StringBuilder(cell.StringCellValue);
            XSSFRichTextString richText = (XSSFRichTextString)cell.RichStringCellValue;
            for (int i = richText.Length - 1; i >= 0; i--)
            {
                IFont f = richText.GetFontAtIndex(i);
                if (f != null && f.IsStrikeout)
                {
                    tmpCellContent.Remove(i, 1);
                }
            }
            String result = tmpCellContent.ToString();
            Assert.AreEqual("320 350", result);
            wb.Close();
        }

        [Test]
        public void Test55406()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("55406_Conditional_formatting_sample.xlsx");
            ISheet sheet = wb.GetSheetAt(0);
            ICell cellA1 = sheet.GetRow(0).GetCell(0);
            ICell cellA2 = sheet.GetRow(1).GetCell(0);

            Assert.AreEqual(0, cellA1.CellStyle.FillForegroundColor);
            Assert.AreEqual("FFFDFDFD", ((XSSFColor)cellA1.CellStyle.FillForegroundColorColor).ARGBHex);
            Assert.AreEqual(0, cellA2.CellStyle.FillForegroundColor);
            Assert.AreEqual("FFFDFDFD", ((XSSFColor)cellA2.CellStyle.FillForegroundColorColor).ARGBHex);

            ISheetConditionalFormatting cond = sheet.SheetConditionalFormatting;
            Assert.AreEqual(2, cond.NumConditionalFormattings);
            Assert.AreEqual(1, cond.GetConditionalFormattingAt(0).NumberOfRules);
            Assert.AreEqual(64, cond.GetConditionalFormattingAt(0).GetRule(0).PatternFormatting.FillForegroundColor);
            Assert.AreEqual("ISEVEN(ROW())", cond.GetConditionalFormattingAt(0).GetRule(0).Formula1);
            Assert.IsNull(((XSSFColor)cond.GetConditionalFormattingAt(0).GetRule(0).PatternFormatting.FillForegroundColorColor).ARGBHex);
            Assert.AreEqual(1, cond.GetConditionalFormattingAt(1).NumberOfRules);
            Assert.AreEqual(64, cond.GetConditionalFormattingAt(1).GetRule(0).PatternFormatting.FillForegroundColor);
            Assert.AreEqual("ISEVEN(ROW())", cond.GetConditionalFormattingAt(1).GetRule(0).Formula1);
            Assert.IsNull(((XSSFColor)cond.GetConditionalFormattingAt(1).GetRule(0).PatternFormatting.FillForegroundColorColor).ARGBHex);

            wb.Close();
        }

        [Test]
        public void Test51998()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("51998.xlsx");

            HashSet<String> sheetNames = new HashSet<String>();

            for (int sheetNum = 0; sheetNum < wb.NumberOfSheets; sheetNum++)
            {
                sheetNames.Add(wb.GetSheetName(sheetNum));
            }

            foreach (String sheetName in sheetNames)
            {
                int sheetIndex = wb.GetSheetIndex(sheetName);

                wb.RemoveSheetAt(sheetIndex);

                ISheet newSheet = wb.CreateSheet();
                //Sheet newSheet = wb.createSheet(sheetName);
                int newSheetIndex = wb.GetSheetIndex(newSheet);
                //System.out.println(newSheetIndex);
                wb.SetSheetName(newSheetIndex, sheetName);
                wb.SetSheetOrder(sheetName, sheetIndex);
            }

            IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            wb.Close();

            Assert.IsNotNull(wbBack);
            wbBack.Close();
        }

        [Test]
        public void Test58731()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("58731.xlsx");
            ISheet sheet = wb.CreateSheet("Java Books");

            Object[][] bookData = {
                new object[] {"Head First Java", "Kathy Serria", 79},
                new object[] {"Effective Java", "Joshua Bloch", 36},
                new object[] {"Clean Code", "Robert martin", 42},
                new object[] {"Thinking in Java", "Bruce Eckel", 35},
        };

            int rowCount = 0;
            foreach (Object[] aBook in bookData)
            {
                IRow row = sheet.CreateRow(rowCount++);

                int columnCount = 0;
                foreach (Object field in aBook)
                {
                    ICell cell = row.CreateCell(columnCount++);
                    if (field is String)
                    {
                        cell.SetCellValue((String)field);
                    }
                    else if (field is int)
                    {
                        cell.SetCellValue((int)field);
                    }
                }
            }

            IWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb2.GetSheet("Java Books");
            Assert.IsNotNull(sheet.GetRow(0));
            Assert.IsNotNull(sheet.GetRow(0).GetCell(0));
            Assert.AreEqual(bookData[0][0], sheet.GetRow(0).GetCell(0).StringCellValue);

            wb2.Close();
            wb.Close();
        }


        /**
         * Regression between 3.10.1 and 3.13 - 
         * org.apache.poi.openxml4j.exceptions.InvalidFormatException: 
         * The part /xl/sharedStrings.xml does not have any content type 
         * ! Rule: Package require content types when retrieving a part from a package. [M.1.14]
         */
        [Test]
        public void Test58760()
        {
            IWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("58760.xlsx");
            Assert.AreEqual(1, wb1.NumberOfSheets);
            Assert.AreEqual("Sheet1", wb1.GetSheetName(0));
            IWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            Assert.AreEqual(1, wb2.NumberOfSheets);
            Assert.AreEqual("Sheet1", wb2.GetSheetName(0));
            wb2.Close();
            wb1.Close();
        }

        [Test]
        public void Test57236()
        {
            // Having very small numbers leads to different formatting, Excel uses the scientific notation, but POI leads to "0"

            /*
            DecimalFormat format = new DecimalFormat("#.##########", new DecimalFormatSymbols(Locale.Default));
            double d = 3.0E-104;
            Assert.AreEqual("3.0E-104", format.Format(d));
             */

            DataFormatter formatter = new DataFormatter(true);

            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57236.xlsx");
            for (int sheetNum = 0; sheetNum < wb.NumberOfSheets; sheetNum++)
            {
                ISheet sheet = wb.GetSheetAt(sheetNum);
                for (int rowNum = sheet.FirstRowNum; rowNum < sheet.LastRowNum; rowNum++)
                {
                    IRow row = sheet.GetRow(rowNum);
                    for (int cellNum = row.FirstCellNum; cellNum < row.LastCellNum; cellNum++)
                    {
                        ICell cell = row.GetCell(cellNum);
                        String fmtCellValue = formatter.FormatCellValue(cell);

                        //System.out.Println("Cell: " + fmtCellValue);
                        Assert.IsNotNull(fmtCellValue);
                        Assert.IsFalse(fmtCellValue.Equals("0"));
                    }
                }
            }
            wb.Close();
        }

        private void CreateXls()
        {
            IWorkbook workbook = new HSSFWorkbook();
            FileStream fileOut = new FileStream("/tmp/rotated.xls", FileMode.Create, FileAccess.ReadWrite);
            ISheet sheet1 = workbook.CreateSheet();
            IRow row1 = sheet1.CreateRow((short)0);
            ICell cell1 = row1.CreateCell(0);
            cell1.SetCellValue("Successful rotated text.");
            ICellStyle style = workbook.CreateCellStyle();
            style.Rotation = ((short)-90);
            cell1.CellStyle = (style);
            workbook.Write(fileOut);
            fileOut.Close();
            workbook.Close();
        }
        private void CreateXlsx()
        {
            IWorkbook workbook = new XSSFWorkbook();
            FileStream fileOut = new FileStream("/tmp/rotated.xlsx", FileMode.Create, FileAccess.ReadWrite);
            ISheet sheet1 = workbook.CreateSheet();
            IRow row1 = sheet1.CreateRow((short)0);
            ICell cell1 = row1.CreateCell(0);
            cell1.SetCellValue("Unsuccessful rotated text.");
            ICellStyle style = workbook.CreateCellStyle();
            style.Rotation = ((short)-90);
            cell1.CellStyle = (style);
            workbook.Write(fileOut);
            fileOut.Close();
            workbook.Close();
        }
        [Ignore("Creates files for checking results manually, actual values are tested in Test*CellStyle")]
        [Test]
        public void Test58043()
        {
            CreateXls();
            CreateXlsx();
        }
        [Test]
        public void Test59132()
        {
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("59132.xlsx");
            ISheet worksheet = workbook.GetSheet("sheet1");
            // B3
            IRow row = worksheet.GetRow(2);
            ICell cell = row.GetCell(1);
            cell.SetCellValue((String)null);
            IFormulaEvaluator evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            // B3
            row = worksheet.GetRow(2);
            cell = row.GetCell(1);
            Assert.AreEqual(CellType.Blank, cell.CellType);
            Assert.AreEqual(CellType.Unknown, evaluator.EvaluateFormulaCell(cell));
            // A3
            row = worksheet.GetRow(2);
            cell = row.GetCell(0);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual("IF(ISBLANK(B3),\"\",B3)", cell.CellFormula);
            Assert.AreEqual(CellType.String, evaluator.EvaluateFormulaCell(cell));
            CellValue value = evaluator.Evaluate(cell);
            Assert.AreEqual("", value.StringValue);
            // A5
            row = worksheet.GetRow(4);
            cell = row.GetCell(0);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual("COUNTBLANK(A1:A4)", cell.CellFormula);
            Assert.AreEqual(CellType.Numeric, evaluator.EvaluateFormulaCell(cell));
            value = evaluator.Evaluate(cell);
            Assert.AreEqual(1.0, value.NumberValue, 0.1);
            /*FileOutputStream output = new FileOutputStream("C:\\temp\\59132.xlsx");
            try {
                workbook.write(output);
            } finally {
                output.close();
            }*/
            workbook.Close();
        }

        [Ignore("bug 59442")]
        [Test]
        public void TestSetRGBBackgroundColor()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFCell cell = workbook.CreateSheet().CreateRow(0).CreateCell(0) as XSSFCell;
            XSSFColor color = new XSSFColor(System.Drawing.Color.Red);
            XSSFCellStyle style = workbook.CreateCellStyle() as XSSFCellStyle;
            style.FillForegroundColorColor = color;
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;
            // Everything is fine at this point, cell is red
            Dictionary<String, Object> properties = new Dictionary<String, Object>();
            properties.Add(CellUtil.BORDER_BOTTOM, BorderStyle.Thin); //or BorderStyle.THIN
            CellUtil.SetCellStyleProperties(cell, properties);

            // Now the cell is all black
            XSSFColor actual = cell.CellStyle.FillBackgroundColorColor as XSSFColor;
            Assert.IsNotNull(actual);
            Assert.AreEqual(color.ARGBHex, actual.ARGBHex);

            XSSFWorkbook nwb = XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            workbook.Close();
            XSSFCell ncell = nwb.GetSheetAt(0).GetRow(0).GetCell(0) as XSSFCell;
            XSSFColor ncolor = new XSSFColor(System.Drawing.Color.Red);
            // Now the cell is all black
            XSSFColor nactual = ncell.CellStyle.FillBackgroundColorColor as XSSFColor;
            Assert.IsNotNull(nactual);
            Assert.AreEqual(ncolor.ARGBHex, nactual.ARGBHex);

            nwb.Close();
        }

        [Ignore("currently Assert.Fails on POI 3.15 beta 2")]
        [Test]
        public void Test55273()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("ExcelTables.xlsx");
            ISheet sheet = wb.GetSheet("ExcelTable");
            IName name = wb.GetName("TableAsRangeName");
            Assert.AreEqual("TableName[#All]", name.RefersToFormula);
            // POI 3.15-beta 2 (2016-06-15): getSheetName throws IllegalArgumentException: Invalid CellReference: TableName[#All]
            Assert.AreEqual("TableName", name.SheetName);
            XSSFSheet xsheet = (XSSFSheet)sheet;
            List<XSSFTable> tables = xsheet.GetTables();
            Assert.AreEqual(2, tables.Count); //FIXME: how many tables are there in this spreadsheet?
            Assert.AreEqual("Table1", tables[0].Name); //FIXME: what is the table name?
            Assert.AreEqual("Table2", tables[1].Name); //FIXME: what is the table name?

            wb.Close();
        }

        [Test]
        public void Test57523()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("57523.xlsx");
            ISheet sheet = wb.GetSheet("Attribute Master");
            IRow row = sheet.GetRow(15);
            int N = CellReference.ConvertColStringToIndex("N");
            ICell N16 = row.GetCell(N);
            Assert.AreEqual(500.0, N16.NumericCellValue, 0.00001);

            int P = CellReference.ConvertColStringToIndex("P");
            ICell P16 = row.GetCell(P);
            Assert.AreEqual(10.0, P16.NumericCellValue, 0.00001);
        }

        /**
         * Files produced by some scientific equipment neglect
         *  to include the row number on the row tags
         */
        [Test]
        public void NoRowNumbers59746()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("59746_NoRowNums.xlsx");
            ISheet sheet = wb.GetSheetAt(0);
            Assert.IsTrue(sheet.LastRowNum > 20, "Last row num: " + sheet.LastRowNum);
            Assert.AreEqual("Checked", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual("Checked", sheet.GetRow(9).GetCell(2).StringCellValue);
            Assert.AreEqual(false, sheet.GetRow(70).GetCell(8).BooleanCellValue);

            Assert.AreEqual(71, sheet.PhysicalNumberOfRows);
            Assert.AreEqual(70, sheet.LastRowNum);
            Assert.AreEqual(70, sheet.GetRow(sheet.LastRowNum).RowNum);
        }

        [Test]
        public void TestWorkdayFunction()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("59106.xlsx");
            XSSFSheet sheet = workbook.GetSheet("Test") as XSSFSheet;
            IRow row = sheet.GetRow(1);
            ICell cell = row.GetCell(0);
            DataFormatter form = new DataFormatter(CultureInfo.GetCultureInfo("en-US"));
            IFormulaEvaluator evaluator = cell.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
            String result = form.FormatCellValue(cell, evaluator);
            Assert.AreEqual("09 Mar 2016", result);
        }
    }

}