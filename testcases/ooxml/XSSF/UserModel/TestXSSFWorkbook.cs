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

using NPOI.XSSF;
using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.Util;
using NPOI.OpenXml4Net.OPC;
using TestCases.HSSF;
using NPOI.XSSF.Model;
using NPOI.OpenXml4Net.OPC.Internal;
using System.Collections.Generic;
using System.Collections;
using System;
using TestCases.SS.UserModel;
using System.Text;
using NPOI.SS.Util;
namespace NPOI.XSSF.UserModel
{

    [TestFixture]
    public class TestXSSFWorkbook : BaseTestWorkbook
    {

        public TestXSSFWorkbook()
            : base(XSSFITestDataProvider.instance)
        {

        }

        /**
         * Tests that we can save, and then re-load a new document
         */
        [Test]
        public void SaveLoadNew()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();

            //check that the default date system is Set to 1900
            CT_WorkbookPr pr = workbook.GetCTWorkbook().workbookPr;
            Assert.IsNotNull(pr);
            Assert.IsTrue(pr.IsSetDate1904());
            Assert.IsFalse(pr.date1904, "XSSF must use the 1900 date system");

            ISheet sheet1 = workbook.CreateSheet("sheet1");
            ISheet sheet2 = workbook.CreateSheet("sheet2");
            workbook.CreateSheet("sheet3");

            IRichTextString rts = workbook.GetCreationHelper().CreateRichTextString("hello world");

            sheet1.CreateRow(0).CreateCell((short)0).SetCellValue(1.2);
            sheet1.CreateRow(1).CreateCell((short)0).SetCellValue(rts);
            sheet2.CreateRow(0);

            Assert.AreEqual(0, workbook.GetSheetAt(0).FirstRowNum);
            Assert.AreEqual(1, workbook.GetSheetAt(0).LastRowNum);
            Assert.AreEqual(0, workbook.GetSheetAt(1).FirstRowNum);
            Assert.AreEqual(0, workbook.GetSheetAt(1).LastRowNum);
            Assert.AreEqual(0, workbook.GetSheetAt(2).FirstRowNum);
            Assert.AreEqual(0, workbook.GetSheetAt(2).LastRowNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream out1 = File.OpenWrite(file.Name);
            workbook.Write(out1);
            out1.Close();

            // Check the namespace Contains what we'd expect it to
            OPCPackage pkg = OPCPackage.Open(file.ToString());
            PackagePart wbRelPart =
                pkg.GetPart(PackagingUriHelper.CreatePartName("/xl/_rels/workbook.xml.rels"));
            Assert.IsNotNull(wbRelPart);
            Assert.IsTrue(wbRelPart.IsRelationshipPart);
            Assert.AreEqual(ContentTypes.RELATIONSHIPS_PART, wbRelPart.ContentType);

            PackagePart wbPart =
                pkg.GetPart(PackagingUriHelper.CreatePartName("/xl/workbook.xml"));
            // Links to the three sheets, shared strings and styles
            Assert.IsTrue(wbPart.HasRelationships);
            Assert.AreEqual(5, wbPart.Relationships.Size);

            // Load back the XSSFWorkbook
            workbook = new XSSFWorkbook(pkg);
            Assert.AreEqual(3, workbook.NumberOfSheets);
            Assert.IsNotNull(workbook.GetSheetAt(0));
            Assert.IsNotNull(workbook.GetSheetAt(1));
            Assert.IsNotNull(workbook.GetSheetAt(2));

            Assert.IsNotNull(workbook.GetSharedStringSource());
            Assert.IsNotNull(workbook.GetStylesSource());

            Assert.AreEqual(0, workbook.GetSheetAt(0).FirstRowNum);
            Assert.AreEqual(1, workbook.GetSheetAt(0).LastRowNum);
            Assert.AreEqual(0, workbook.GetSheetAt(1).FirstRowNum);
            Assert.AreEqual(0, workbook.GetSheetAt(1).LastRowNum);
            Assert.AreEqual(0, workbook.GetSheetAt(2).FirstRowNum);
            Assert.AreEqual(0, workbook.GetSheetAt(2).LastRowNum);

            sheet1 = workbook.GetSheetAt(0);
            Assert.AreEqual(1.2, sheet1.GetRow(0).GetCell(0).NumericCellValue, 0.0001);
            Assert.AreEqual("hello world", sheet1.GetRow(1).GetCell(0).RichStringCellValue.String);

            pkg.Close();
        }
        [Test]
        public void Existing()
        {

            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("Formatting.xlsx");
            Assert.IsNotNull(workbook.GetSharedStringSource());
            Assert.IsNotNull(workbook.GetStylesSource());

            // And check a few low level bits too
            OPCPackage pkg = OPCPackage.Open(HSSFTestDataSamples.OpenSampleFileStream("Formatting.xlsx"));
            PackagePart wbPart =
                pkg.GetPart(PackagingUriHelper.CreatePartName("/xl/workbook.xml"));

            // Links to the three sheets, shared, styles and themes
            Assert.IsTrue(wbPart.HasRelationships);
            Assert.AreEqual(6, wbPart.Relationships.Size);
            pkg.Close();
        }
        [Test]
        public void GetCellStyleAt()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            short i = 0;
            //get default style
            ICellStyle cellStyleAt = workbook.GetCellStyleAt(i);
            Assert.IsNotNull(cellStyleAt);

            //get custom style
            StylesTable styleSource = workbook.GetStylesSource();
            XSSFCellStyle customStyle = new XSSFCellStyle(styleSource);
            XSSFFont font = new XSSFFont();
            font.FontName = ("Verdana");
            customStyle.SetFont(font);
            int x = styleSource.PutStyle(customStyle);
            cellStyleAt = workbook.GetCellStyleAt((short)x);
            Assert.IsNotNull(cellStyleAt);
        }
        [Test]
        public void GetFontAt()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            StylesTable styleSource = workbook.GetStylesSource();
            short i = 0;
            //get default font
            IFont fontAt = workbook.GetFontAt(i);
            Assert.IsNotNull(fontAt);

            //get customized font
            XSSFFont customFont = new XSSFFont();
            customFont.IsItalic = (true);
            int x = styleSource.PutFont(customFont);
            fontAt = workbook.GetFontAt((short)x);
            Assert.IsNotNull(fontAt);
        }
        [Test]
        public void NumCellStyles()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            short i = workbook.NumCellStyles;
            //get default cellStyles
            Assert.AreEqual(1, i);
        }
        [Test]
        public void LoadSave()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("Formatting.xlsx");
            Assert.AreEqual(3, workbook.NumberOfSheets);
            Assert.AreEqual("dd/mm/yyyy", workbook.GetSheetAt(0).GetRow(1).GetCell(0).RichStringCellValue.String);
            Assert.IsNotNull(workbook.GetSharedStringSource());
            Assert.IsNotNull(workbook.GetStylesSource());

            // Write out, and check
            // Load up again, check all still there
            XSSFWorkbook wb2 = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            Assert.AreEqual(3, wb2.NumberOfSheets);
            Assert.IsNotNull(wb2.GetSheetAt(0));
            Assert.IsNotNull(wb2.GetSheetAt(1));
            Assert.IsNotNull(wb2.GetSheetAt(2));

            Assert.AreEqual("dd/mm/yyyy", wb2.GetSheetAt(0).GetRow(1).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("yyyy/mm/dd", wb2.GetSheetAt(0).GetRow(2).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("yyyy-mm-dd", wb2.GetSheetAt(0).GetRow(3).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("yy/mm/dd", wb2.GetSheetAt(0).GetRow(4).GetCell(0).RichStringCellValue.String);
            Assert.IsNotNull(wb2.GetSharedStringSource());
            Assert.IsNotNull(wb2.GetStylesSource());
        }
        [Test]
        public void Styles()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("Formatting.xlsx");

            StylesTable ss = workbook.GetStylesSource();
            Assert.IsNotNull(ss);
            StylesTable st = ss;

            // Has 8 number formats
            Assert.AreEqual(8, st.NumberFormatSize);
            // Has 2 fonts
            Assert.AreEqual(2, st.GetFonts().Count);
            // Has 2 Fills
            Assert.AreEqual(2, st.GetFills().Count);
            // Has 1 border
            Assert.AreEqual(1, st.GetBorders().Count);

            // Add two more styles
            Assert.AreEqual(StylesTable.FIRST_CUSTOM_STYLE_ID + 8,
                    st.PutNumberFormat("testFORMAT"));
            Assert.AreEqual(StylesTable.FIRST_CUSTOM_STYLE_ID + 8,
                    st.PutNumberFormat("testFORMAT"));
            Assert.AreEqual(StylesTable.FIRST_CUSTOM_STYLE_ID + 9,
                    st.PutNumberFormat("testFORMAT2"));
            Assert.AreEqual(10, st.NumberFormatSize);


            // Save, load back in again, and check
            workbook = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);

            ss = workbook.GetStylesSource();
            Assert.IsNotNull(ss);

            Assert.AreEqual(10, st.NumberFormatSize);
            Assert.AreEqual(2, st.GetFonts().Count);
            Assert.AreEqual(2, st.GetFills().Count);
            Assert.AreEqual(1, st.GetBorders().Count);
        }
        [Test]
        public void IncrementSheetId()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                int sheetId = (int)(wb.CreateSheet() as XSSFSheet).sheet.sheetId;
                Assert.AreEqual(1, sheetId);
                sheetId = (int)(wb.CreateSheet() as XSSFSheet).sheet.sheetId;
                Assert.AreEqual(2, sheetId);

                //test file with gaps in the sheetId sequence
                XSSFWorkbook wbBack = XSSFTestDataSamples.OpenSampleWorkbook("47089.xlsm");
                try
                {
                    int lastSheetId = (int)(wbBack.GetSheetAt(wbBack.NumberOfSheets - 1)  as XSSFSheet).sheet.sheetId;
                    sheetId = (int)(wbBack.CreateSheet() as XSSFSheet).sheet.sheetId;
                    Assert.AreEqual(lastSheetId + 1, sheetId);
                }
                finally
                {
                    wbBack.Close();
                }
            }
            finally
            {
                wb.Close();
            }
        }

        /**
         *  Test Setting of core properties such as Title and Author
         */
        [Test]
        public void WorkbookProperties()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            try
            {
                POIXMLProperties props = workbook.GetProperties();
                Assert.IsNotNull(props);
                //the Application property must be set for new workbooks, see Bugzilla #47559
                Assert.AreEqual("NPOI", props.ExtendedProperties.GetUnderlyingProperties().Application);

                PackagePropertiesPart opcProps = props.CoreProperties.GetUnderlyingProperties();
                Assert.IsNotNull(opcProps);

                opcProps.SetTitleProperty("Testing Bugzilla #47460");
                Assert.AreEqual("NPOI", opcProps.GetCreatorProperty());
                opcProps.SetCreatorProperty("poi-dev@poi.apache.org");

                XSSFWorkbook wbBack = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
                Assert.AreEqual("NPOI", wbBack.GetProperties().ExtendedProperties.GetUnderlyingProperties().Application);
                opcProps = wbBack.GetProperties().CoreProperties.GetUnderlyingProperties();
                Assert.AreEqual("Testing Bugzilla #47460", opcProps.GetTitleProperty());
                Assert.AreEqual("poi-dev@poi.apache.org", opcProps.GetCreatorProperty());
            }
            finally
            {
                workbook.Close();
            }
        }

        /**
         * Verify that the attached Test data was not modified. If this Test method
         * fails, the Test data is not working properly.
         */
        //public void Bug47668()
        //{
        //    XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("47668.xlsx");
        //    IList allPictures = workbook.GetAllPictures();
        //    Assert.AreEqual(1, allPictures.Count);

        //    PackagePartName imagePartName = PackagingUriHelper
        //            .CreatePartName("/xl/media/image1.jpeg");
        //    PackagePart imagePart = workbook.Package.GetPart(imagePartName);
        //    Assert.IsNotNull(imagePart);

        //    foreach (XSSFPictureData pictureData in allPictures)
        //    {
        //        PackagePart picturePart = pictureData.GetPackagePart();
        //        Assert.AreSame(imagePart, picturePart);
        //    }

        //    XSSFSheet sheet0 = (XSSFSheet)workbook.GetSheetAt(0);
        //    XSSFDrawing Drawing0 = (XSSFDrawing)sheet0.CreateDrawingPatriarch();
        //    XSSFPictureData pictureData0 = (XSSFPictureData)Drawing0.GetRelations()[0];
        //    byte[] data0 = pictureData0.Data;
        //    CRC32 crc0 = new CRC32();
        //    crc0.Update(data0);

        //    XSSFSheet sheet1 = workbook.GetSheetAt(1);
        //    XSSFDrawing Drawing1 = sheet1.CreateDrawingPatriarch();
        //    XSSFPictureData pictureData1 = (XSSFPictureData)Drawing1.GetRelations()[0];
        //    byte[] data1 = pictureData1.Data;
        //    CRC32 crc1 = new CRC32();
        //    crc1.Update(data1);

        //    Assert.AreEqual(crc0.GetValue(), crc1.GetValue());
        //}

        /**
         * When deleting a sheet make sure that we adjust sheet indices of named ranges
         */
        [Test]
        public void Bug47737()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("47737.xlsx");
            Assert.AreEqual(2, wb.NumberOfNames);
            Assert.IsNotNull(wb.GetCalculationChain());

            XSSFName nm0 = (XSSFName)wb.GetNameAt(0);
            Assert.IsTrue(nm0.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual(0u, nm0.GetCTName().localSheetId);

            XSSFName nm1 = (XSSFName)wb.GetNameAt(1);
            Assert.IsTrue(nm1.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual(1u, nm1.GetCTName().localSheetId);

            wb.RemoveSheetAt(0);
            Assert.AreEqual(1, wb.NumberOfNames);
            XSSFName nm2 = (XSSFName)wb.GetNameAt(0);
            Assert.IsTrue(nm2.GetCTName().IsSetLocalSheetId());
            Assert.AreEqual(0u, nm2.GetCTName().localSheetId);
            //calculation chain is Removed as well
            Assert.IsNull(wb.GetCalculationChain());

        }

        /**
         * Problems with XSSFWorkbook.RemoveSheetAt when workbook Contains chart
         */
        [Test]
        public void Bug47813()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("47813.xlsx");
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.IsNotNull(wb.GetCalculationChain());

            Assert.AreEqual("Numbers", wb.GetSheetName(0));
            //the second sheet is of type 'chartsheet'
            Assert.AreEqual("Chart", wb.GetSheetName(1));
            Assert.IsTrue(wb.GetSheetAt(1) is XSSFChartSheet);
            Assert.AreEqual("SomeJunk", wb.GetSheetName(2));

            wb.RemoveSheetAt(2);
            Assert.AreEqual(2, wb.NumberOfSheets);
            Assert.IsNull(wb.GetCalculationChain());

            wb = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.AreEqual(2, wb.NumberOfSheets);
            Assert.IsNull(wb.GetCalculationChain());

            Assert.AreEqual("Numbers", wb.GetSheetName(0));
            Assert.AreEqual("Chart", wb.GetSheetName(1));
        }

        /**
         * Problems with the count of the number of styles
         *  coming out wrong
         */
        [Test]
        public void Bug49702()
        {
            // First try with a new file
            XSSFWorkbook wb = new XSSFWorkbook();

            // Should have one style
            Assert.AreEqual(1, wb.NumCellStyles);
            wb.GetCellStyleAt((short)0);
            try
            {
                wb.GetCellStyleAt((short)1);
                Assert.Fail("Shouldn't be able to get style at 1 that doesn't exist");
            }
            catch (ArgumentOutOfRangeException) { }

            // Add another one
            ICellStyle cs = wb.CreateCellStyle();
            cs.DataFormat = ((short)11);

            // Re-check
            Assert.AreEqual(2, wb.NumCellStyles);
            wb.GetCellStyleAt((short)0);
            wb.GetCellStyleAt((short)1);
            try
            {
                wb.GetCellStyleAt((short)2);
                Assert.Fail("Shouldn't be able to get style at 2 that doesn't exist");
            }
            catch (ArgumentOutOfRangeException) { }

            // Save and reload
            XSSFWorkbook nwb = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.AreEqual(2, nwb.NumCellStyles);
            nwb.GetCellStyleAt((short)0);
            nwb.GetCellStyleAt((short)1);
            try
            {
                nwb.GetCellStyleAt((short)2);
                Assert.Fail("Shouldn't be able to Get style at 2 that doesn't exist");
            }
            catch (ArgumentOutOfRangeException) { }

            // Now with an existing file
            wb = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            Assert.AreEqual(3, wb.NumCellStyles);
            wb.GetCellStyleAt((short)0);
            wb.GetCellStyleAt((short)1);
            wb.GetCellStyleAt((short)2);
            try
            {
                wb.GetCellStyleAt((short)3);
                Assert.Fail("Shouldn't be able to Get style at 3 that doesn't exist");
            }
            catch (ArgumentOutOfRangeException) { }
        }
        [Test]
        public void RecalcId()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            Assert.IsFalse(wb.GetForceFormulaRecalculation());
            CT_Workbook ctWorkbook = wb.GetCTWorkbook();
            Assert.IsFalse(ctWorkbook.IsSetCalcPr());

            wb.SetForceFormulaRecalculation(true); // resets the EngineId flag to zero

            CT_CalcPr calcPr = ctWorkbook.calcPr;
            Assert.IsNotNull(calcPr);
            Assert.AreEqual(0, (int)calcPr.calcId);

            calcPr.calcId = 100;
            Assert.IsTrue(wb.GetForceFormulaRecalculation());

            wb.SetForceFormulaRecalculation(true); // resets the EngineId flag to zero
            Assert.AreEqual(0, (int)calcPr.calcId);
            Assert.IsFalse(wb.GetForceFormulaRecalculation());

            // calcMode="manual" is unset when forceFormulaRecalculation=true
            calcPr.calcMode = (ST_CalcMode.manual);
            wb.SetForceFormulaRecalculation(true);
            Assert.AreEqual(ST_CalcMode.auto, calcPr.calcMode);
        }
        [Test]
        public void ChangeSheetNameWithSharedFormulas()
        {
            ChangeSheetNameWithSharedFormulas("shared_formulas.xlsx");
        }
        [Test]
        public void SetTabColor()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sh = wb.CreateSheet() as XSSFSheet;
            Assert.IsTrue(sh.GetCTWorksheet().sheetPr == null || !sh.GetCTWorksheet().sheetPr.IsSetTabColor());
            sh.SetTabColor(IndexedColors.Red.Index);
            Assert.IsTrue(sh.GetCTWorksheet().sheetPr.IsSetTabColor());
            Assert.AreEqual(IndexedColors.Red.Index,
                    sh.GetCTWorksheet().sheetPr.tabColor.indexed);
        }
        [Test]
        public void ColumnWidthPOI52233()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("hello world");

            sheet = workbook.CreateSheet();
            sheet.SetColumnWidth(4, 5000);
            sheet.SetColumnWidth(5, 5000);

            sheet.GroupColumn((short)4, (short)5);

            accessWorkbook(workbook);

            MemoryStream stream = new MemoryStream();
            try
            {
                workbook.Write(stream);
            }
            finally
            {
                stream.Close();
            }
            accessWorkbook(workbook);

        }

        private void accessWorkbook(XSSFWorkbook workbook)
        {
            workbook.GetSheetAt(1).SetColumnGroupCollapsed(4, true);
            workbook.GetSheetAt(1).SetColumnGroupCollapsed(4, false);

            Assert.AreEqual("hello world", workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual(2048, workbook.GetSheetAt(0).GetColumnWidth(0)); // <-works
        }

        [Test]
        public void Bug48495()
        {
            try
            {
                IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48495.xlsx");

                assertSheetOrder(wb, "Sheet1");

                ISheet sheet = wb.GetSheetAt(0);
                sheet.ShiftRows(2, sheet.LastRowNum, 1, true, false);
                IRow newRow = sheet.GetRow(2);
                if (newRow == null) newRow = sheet.CreateRow(2);
                newRow.CreateCell(0).SetCellValue(" Another Header");
                wb.CloneSheet(0);

                assertSheetOrder(wb, "Sheet1", "Sheet1 (2)");


                IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(wb);
                Assert.IsNotNull(read);
                assertSheetOrder(read, "Sheet1", "Sheet1 (2)");
            }
            catch (Exception e)
            {
            }

        }
        [Test]
        public void Bug47090a()
        {
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("47090.xlsx");
            assertSheetOrder(workbook, "Sheet1", "Sheet2");
            workbook.RemoveSheetAt(0);
            assertSheetOrder(workbook, "Sheet2");
            workbook.CreateSheet();
            assertSheetOrder(workbook, "Sheet2", "Sheet1");
            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            assertSheetOrder(read, "Sheet2", "Sheet1");
        }

        [Test]
        public void Bug47090b()
        {
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("47090.xlsx");
            assertSheetOrder(workbook, "Sheet1", "Sheet2");
            workbook.RemoveSheetAt(1);
            assertSheetOrder(workbook, "Sheet1");
            workbook.CreateSheet();
            assertSheetOrder(workbook, "Sheet1", "Sheet0");		// Sheet0 because it uses "Sheet" + sheets.size() as starting point!
            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            assertSheetOrder(read, "Sheet1", "Sheet0");
        }

        [Test]
        public void Bug47090c()
        {
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("47090.xlsx");
            assertSheetOrder(workbook, "Sheet1", "Sheet2");
            workbook.RemoveSheetAt(0);
            assertSheetOrder(workbook, "Sheet2");
            workbook.CloneSheet(0);
            assertSheetOrder(workbook, "Sheet2", "Sheet2 (2)");
            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            assertSheetOrder(read, "Sheet2", "Sheet2 (2)");
        }

        [Test]
        public void Bug47090d()
        {
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("47090.xlsx");
            assertSheetOrder(workbook, "Sheet1", "Sheet2");
            workbook.CreateSheet();
            assertSheetOrder(workbook, "Sheet1", "Sheet2", "Sheet0");
            workbook.RemoveSheetAt(0);
            assertSheetOrder(workbook, "Sheet2", "Sheet0");
            workbook.CreateSheet();
            assertSheetOrder(workbook, "Sheet2", "Sheet0", "Sheet1");
            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            assertSheetOrder(read, "Sheet2", "Sheet0", "Sheet1");
        }

        [Test]
        public void Bug51158()
        {
            // create a workbook
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.CreateSheet("Test Sheet") as XSSFSheet;
            XSSFRow row = sheet.CreateRow(2) as XSSFRow;
            XSSFCell cell = row.CreateCell(3) as XSSFCell;
            cell.SetCellValue("test1");

            //XSSFCreationHelper helper = workbook.GetCreationHelper();
            //cell.Hyperlink=(/*setter*/helper.CreateHyperlink(0));

            XSSFComment comment = (sheet.CreateDrawingPatriarch() as XSSFDrawing).CreateCellComment(new XSSFClientAnchor()) as XSSFComment;
            Assert.IsNotNull(comment);
            comment.SetString("some comment");

            //        ICellStyle cs = workbook.CreateCellStyle();
            //        cs.ShrinkToFit=(/*setter*/false);
            //        row.CreateCell(0).CellStyle=(/*setter*/cs);

            // write the first excel file
            XSSFWorkbook readBack = XSSFTestDataSamples.WriteOutAndReadBack(workbook) as XSSFWorkbook;
            Assert.IsNotNull(readBack);
            Assert.AreEqual("test1", readBack.GetSheetAt(0).GetRow(2).GetCell(3).StringCellValue);
            Assert.IsNull(readBack.GetSheetAt(0).GetRow(2).GetCell(4));

            // add a new cell to the sheet
            cell = row.CreateCell(4) as XSSFCell;
            cell.SetCellValue("test2");

            // write the second excel file
            readBack = XSSFTestDataSamples.WriteOutAndReadBack(workbook) as XSSFWorkbook;
            Assert.IsNotNull(readBack);
            Assert.AreEqual("test1", readBack.GetSheetAt(0).GetRow(2).GetCell(3).StringCellValue);
            Assert.AreEqual("test2", readBack.GetSheetAt(0).GetRow(2).GetCell(4).StringCellValue);
        }

        [Test]
        public void Bug51158a()
        {
            // create a workbook
            XSSFWorkbook workbook = new XSSFWorkbook();
            workbook.CreateSheet("Test Sheet");

            XSSFSheet sheetBack = workbook.GetSheetAt(0) as XSSFSheet;

            // Committing twice did add the XML twice without Clearing the part in between
            sheetBack.Commit();

            // ensure that a memory based package part does not have lingering data from previous Commit() calls
            if (sheetBack.GetPackagePart() is MemoryPackagePart)
            {
                ((MemoryPackagePart)sheetBack.GetPackagePart()).Clear();
            }

            sheetBack.Commit();

            String str = Encoding.UTF8.GetString(IOUtils.ToByteArray(sheetBack.GetPackagePart().GetInputStream()));
            //System.out.Println(str);

            Assert.AreEqual(1, countMatches(str, "<worksheet"));
        }

        private static int INDEX_NOT_FOUND = -1;

        private static int countMatches(string str, string sub)
        {
            if (string.IsNullOrEmpty(str) || string.IsNullOrEmpty(sub))
            {
                return 0;
            }
            int count = 0;
            int idx = 0;
            while ((idx = IndexOf(str, sub, idx)) != INDEX_NOT_FOUND)
            {
                count++;
                idx += sub.Length;
            }
            return count;
        }

        private static int IndexOf(string cs, string searchChar, int start)
        {
            return cs.ToString().IndexOf(searchChar.ToString(), start);
        }

        [Test]
        public void AddPivotCache()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                CT_Workbook ctWb = wb.GetCTWorkbook();
                CT_PivotCache pivotCache = wb.AddPivotCache("0");
                //Ensures that pivotCaches is Initiated
                Assert.IsTrue(ctWb.IsSetPivotCaches());
                Assert.AreSame(pivotCache, ctWb.pivotCaches.GetPivotCacheArray(0));
                Assert.AreEqual("0", pivotCache.id);
            }
            finally
            {
                wb.Close();
            }
        }

        public void SetPivotData(XSSFWorkbook wb)
        {
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

            IRow row1 = sheet.CreateRow(0);
            // Create a cell and Put a value in it.
            ICell cell = row1.CreateCell(0);
            cell.SetCellValue("Names");
            ICell cell2 = row1.CreateCell(1);
            cell2.SetCellValue("#");
            ICell cell7 = row1.CreateCell(2);
            cell7.SetCellValue("Data");

            IRow row2 = sheet.CreateRow(1);
            ICell cell3 = row2.CreateCell(0);
            cell3.SetCellValue("Jan");
            ICell cell4 = row2.CreateCell(1);
            cell4.SetCellValue(10);
            ICell cell8 = row2.CreateCell(2);
            cell8.SetCellValue("Apa");

            IRow row3 = sheet.CreateRow(2);
            ICell cell5 = row3.CreateCell(0);
            cell5.SetCellValue("Ben");
            ICell cell6 = row3.CreateCell(1);
            cell6.SetCellValue(9);
            ICell cell9 = row3.CreateCell(2);
            cell9.SetCellValue("Bepa");

            AreaReference source = new AreaReference("A1:B2");
            sheet.CreatePivotTable(source, new CellReference("H5"));
        }

        [Test]
        public void LoadWorkbookWithPivotTable()
        {
            String fileName = "ooxml-pivottable.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook();
            SetPivotData(wb);

            FileStream fileOut = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
            wb.Write(fileOut);
            fileOut.Close();

            XSSFWorkbook wb2 = (XSSFWorkbook)WorkbookFactory.Create(fileName);
            Assert.IsTrue(wb2.PivotTables.Count == 1);
        }

        [Test]
        public void AddPivotTableToWorkbookWithLoadedPivotTable()
        {
            String fileName = "ooxml-pivottable.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook();
            SetPivotData(wb);

            FileStream fileOut = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
            wb.Write(fileOut);
            fileOut.Close();

            XSSFWorkbook wb2 = (XSSFWorkbook)WorkbookFactory.Create(fileName);
            SetPivotData(wb2);
            Assert.IsTrue(wb2.PivotTables.Count == 2);
        }

        [Test]
        public void SetFirstVisibleTab_57373()
        {
            XSSFWorkbook wb = new XSSFWorkbook();

            try
            {
                /*Sheet sheet1 =*/
                wb.CreateSheet();
                ISheet sheet2 = wb.CreateSheet();
                int idx2 = wb.GetSheetIndex(sheet2);
                ISheet sheet3 = wb.CreateSheet();
                int idx3 = wb.GetSheetIndex(sheet3);

                // add many sheets so "first visible" is relevant
                for (int i = 0; i < 30; i++)
                {
                    wb.CreateSheet();
                }

                wb.FirstVisibleTab = (/*setter*/idx2);
                wb.SetActiveSheet(idx3);

                //wb.Write(new FileOutputStream(new File("C:\\temp\\test.xlsx")));

                Assert.AreEqual(idx2, wb.FirstVisibleTab);
                Assert.AreEqual(idx3, wb.ActiveSheetIndex);

                IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);

                sheet2 = wbBack.GetSheetAt(idx2);
                sheet3 = wbBack.GetSheetAt(idx3);
                Assert.AreEqual(idx2, wb.FirstVisibleTab);
                Assert.AreEqual(idx3, wb.ActiveSheetIndex);
            }
            finally
            {
                wb.Close();
            }
        }

    }
}
