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

using NPOI;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using NUnit.Framework;using NUnit.Framework.Legacy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TestCases.HSSF;
using TestCases.SS.UserModel;

namespace TestCases.XSSF.UserModel
{

    [TestFixture]
    public class TestXSSFWorkbook : BaseTestXWorkbook
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
            XSSFWorkbook wb1 = new XSSFWorkbook();

            //check that the default date system is Set to 1900
            CT_WorkbookPr pr = wb1.GetCTWorkbook().workbookPr;
            ClassicAssert.IsNotNull(pr);
            ClassicAssert.IsTrue(pr.IsSetDate1904());
            ClassicAssert.IsFalse(pr.date1904, "XSSF must use the 1900 date system");

            ISheet sheet1 = wb1.CreateSheet("sheet1");
            ISheet sheet2 = wb1.CreateSheet("sheet2");
            wb1.CreateSheet("sheet3");
            ISheet sheet4 = wb1.CreateSheet("sheet4");

            IRichTextString rts = wb1.GetCreationHelper().CreateRichTextString("hello world");

            sheet1.CreateRow(0).CreateCell((short)0).SetCellValue(1.2);
            sheet1.CreateRow(1).CreateCell((short)0).SetCellValue(rts);
            sheet2.CreateRow(0);
            ((XSSFSheet)sheet2).CreateColumn(0);

            ((XSSFSheet)sheet4).CreateColumn(0).CreateCell(0).SetCellValue(1.2);
            ((XSSFSheet)sheet4).CreateColumn(1).CreateCell(0).SetCellValue(rts);

            ClassicAssert.AreEqual(0, wb1.GetSheetAt(0).FirstRowNum);
            ClassicAssert.AreEqual(1, wb1.GetSheetAt(0).LastRowNum);
            ClassicAssert.AreEqual(0, wb1.GetSheetAt(1).FirstRowNum);
            ClassicAssert.AreEqual(0, wb1.GetSheetAt(1).LastRowNum);
            ClassicAssert.AreEqual(0, wb1.GetSheetAt(2).FirstRowNum);
            ClassicAssert.AreEqual(0, wb1.GetSheetAt(2).LastRowNum);
            ClassicAssert.AreEqual(0, wb1.GetSheetAt(3).FirstRowNum);
            ClassicAssert.AreEqual(0, wb1.GetSheetAt(3).LastRowNum);

            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(0)).FirstColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(0)).LastColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(1)).FirstColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(1)).LastColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(2)).FirstColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(2)).LastColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(3)).FirstColumnNum);
            ClassicAssert.AreEqual(1, ((XSSFSheet)wb1.GetSheetAt(3)).LastColumnNum);

            FileInfo file = TempFile.CreateTempFile("poi-", ".xlsx");
            Stream out1 = File.OpenWrite(file.FullName);
            wb1.Write(out1);
            out1.Close();

            // Check the namespace Contains what we'd expect it to
            OPCPackage pkg = OPCPackage.Open(file.ToString());
            PackagePart wbRelPart =
                pkg.GetPart(PackagingUriHelper.CreatePartName("/xl/_rels/workbook.xml.rels"));
            ClassicAssert.IsNotNull(wbRelPart);
            ClassicAssert.IsTrue(wbRelPart.IsRelationshipPart);
            ClassicAssert.AreEqual(ContentTypes.RELATIONSHIPS_PART, wbRelPart.ContentType);

            PackagePart wbPart =
                pkg.GetPart(PackagingUriHelper.CreatePartName("/xl/workbook.xml"));
            // Links to the three sheets, shared strings and styles
            ClassicAssert.IsTrue(wbPart.HasRelationships);
            ClassicAssert.AreEqual(6, wbPart.Relationships.Size);
            wb1.Close();

            // Load back the XSSFWorkbook
            XSSFWorkbook wb2 = new XSSFWorkbook(pkg);
            ClassicAssert.AreEqual(4, wb2.NumberOfSheets);
            ClassicAssert.IsNotNull(wb2.GetSheetAt(0));
            ClassicAssert.IsNotNull(wb2.GetSheetAt(1));
            ClassicAssert.IsNotNull(wb2.GetSheetAt(2));
            ClassicAssert.IsNotNull(wb2.GetSheetAt(3));

            ClassicAssert.IsNotNull(wb2.GetSharedStringSource());
            ClassicAssert.IsNotNull(wb2.GetStylesSource());

            ClassicAssert.AreEqual(0, wb2.GetSheetAt(0).FirstRowNum);
            ClassicAssert.AreEqual(1, wb2.GetSheetAt(0).LastRowNum);
            ClassicAssert.AreEqual(0, wb2.GetSheetAt(1).FirstRowNum);
            ClassicAssert.AreEqual(0, wb2.GetSheetAt(1).LastRowNum);
            ClassicAssert.AreEqual(0, wb2.GetSheetAt(2).FirstRowNum);
            ClassicAssert.AreEqual(0, wb2.GetSheetAt(2).LastRowNum);
            ClassicAssert.AreEqual(0, wb2.GetSheetAt(3).FirstRowNum);
            ClassicAssert.AreEqual(0, wb2.GetSheetAt(3).LastRowNum);

            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(0)).FirstColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(0)).LastColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(1)).FirstColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(1)).LastColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(2)).FirstColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(2)).LastColumnNum);
            ClassicAssert.AreEqual(0, ((XSSFSheet)wb1.GetSheetAt(3)).FirstColumnNum);
            ClassicAssert.AreEqual(1, ((XSSFSheet)wb1.GetSheetAt(3)).LastColumnNum);

            sheet1 = wb2.GetSheetAt(0);
            ClassicAssert.AreEqual(1.2, sheet1.GetRow(0).GetCell(0).NumericCellValue, 0.0001);
            ClassicAssert.AreEqual(1.2, ((XSSFSheet)sheet4).GetColumn(0).GetCell(0).NumericCellValue, 0.0001);
            ClassicAssert.AreEqual("hello world", sheet1.GetRow(1).GetCell(0).RichStringCellValue.String);
            ClassicAssert.AreEqual("hello world", ((XSSFSheet)sheet4).GetColumn(1).GetCell(0).RichStringCellValue.String);

            pkg.Close();
        }

        [Test]
        public void Existing()
        {

            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("Formatting.xlsx");
            ClassicAssert.IsNotNull(workbook.GetSharedStringSource());
            ClassicAssert.IsNotNull(workbook.GetStylesSource());

            // And check a few low level bits too
            OPCPackage pkg = OPCPackage.Open(HSSFTestDataSamples.OpenSampleFileStream("Formatting.xlsx"));
            PackagePart wbPart =
                pkg.GetPart(PackagingUriHelper.CreatePartName("/xl/workbook.xml"));

            // Links to the three sheets, shared, styles and themes
            ClassicAssert.IsTrue(wbPart.HasRelationships);
            ClassicAssert.AreEqual(6, wbPart.Relationships.Size);
            pkg.Close();
            workbook.Close();
        }
        [Test]
        public void GetCellStyleAt()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            try
            {
                short i = 0;
                //get default style
                ICellStyle cellStyleAt = workbook.GetCellStyleAt(i);
                ClassicAssert.IsNotNull(cellStyleAt);

                //get custom style
                StylesTable styleSource = workbook.GetStylesSource();
                XSSFCellStyle customStyle = new XSSFCellStyle(styleSource);
                XSSFFont font = new XSSFFont();
                font.FontName = ("Verdana");
                customStyle.SetFont(font);
                int x = styleSource.PutStyle(customStyle);
                cellStyleAt = workbook.GetCellStyleAt((short)x);
                ClassicAssert.IsNotNull(cellStyleAt);
            }
            finally
            {

            }

        }
        [Test]
        public void GetFontAt()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            try
            {
                StylesTable styleSource = workbook.GetStylesSource();
                short i = 0;
                //get default font
                IFont fontAt = workbook.GetFontAt(i);
                ClassicAssert.IsNotNull(fontAt);

                //get customized font
                XSSFFont customFont = new XSSFFont();
                customFont.IsItalic = (true);
                int x = styleSource.PutFont(customFont);
                fontAt = workbook.GetFontAt((short)x);
                ClassicAssert.IsNotNull(fontAt);
            }
            finally
            {
                workbook.Close();
            }

        }
        [Test]
        public void GetNumCellStyles()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            try
            {
                //get default cellStyles
                ClassicAssert.AreEqual(1, workbook.NumCellStyles);
            }
            finally
            {
                workbook.Close();
            }
        }
        [Test]
        public void LoadSave()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("Formatting.xlsx");
            ClassicAssert.AreEqual(3, workbook.NumberOfSheets);
            ClassicAssert.AreEqual("dd/mm/yyyy", workbook.GetSheetAt(0).GetRow(1).GetCell(0).RichStringCellValue.String);
            ClassicAssert.IsNotNull(workbook.GetSharedStringSource());
            ClassicAssert.IsNotNull(workbook.GetStylesSource());

            // Write out, and check
            // Load up again, check all still there
            XSSFWorkbook wb2 = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            ClassicAssert.AreEqual(3, wb2.NumberOfSheets);
            ClassicAssert.IsNotNull(wb2.GetSheetAt(0));
            ClassicAssert.IsNotNull(wb2.GetSheetAt(1));
            ClassicAssert.IsNotNull(wb2.GetSheetAt(2));

            ClassicAssert.AreEqual("dd/mm/yyyy", wb2.GetSheetAt(0).GetRow(1).GetCell(0).RichStringCellValue.String);
            ClassicAssert.AreEqual("yyyy/mm/dd", wb2.GetSheetAt(0).GetRow(2).GetCell(0).RichStringCellValue.String);
            ClassicAssert.AreEqual("yyyy-mm-dd", wb2.GetSheetAt(0).GetRow(3).GetCell(0).RichStringCellValue.String);
            ClassicAssert.AreEqual("yy/mm/dd", wb2.GetSheetAt(0).GetRow(4).GetCell(0).RichStringCellValue.String);
            ClassicAssert.IsNotNull(wb2.GetSharedStringSource());
            ClassicAssert.IsNotNull(wb2.GetStylesSource());

            workbook.Close();
            wb2.Close();
        }
        [Test]
        public void Styles()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("Formatting.xlsx");

            StylesTable ss = wb1.GetStylesSource();
            ClassicAssert.IsNotNull(ss);
            StylesTable st = ss;

            // Has 8 number formats
            ClassicAssert.AreEqual(8, st.NumDataFormats);
            // Has 2 fonts
            ClassicAssert.AreEqual(2, st.Fonts.Count);
            // Has 2 Fills
            ClassicAssert.AreEqual(2, st.GetFills().Count);
            // Has 1 border
            ClassicAssert.AreEqual(1, st.GetBorders().Count);

            // Add two more styles
            ClassicAssert.AreEqual(StylesTable.FIRST_CUSTOM_STYLE_ID + 8,
                    st.PutNumberFormat("testFORMAT"));
            ClassicAssert.AreEqual(StylesTable.FIRST_CUSTOM_STYLE_ID + 8,
                    st.PutNumberFormat("testFORMAT"));
            ClassicAssert.AreEqual(StylesTable.FIRST_CUSTOM_STYLE_ID + 9,
                    st.PutNumberFormat("testFORMAT2"));
            ClassicAssert.AreEqual(10, st.NumDataFormats);


            // Save, load back in again, and check
            XSSFWorkbook wb2 = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            ss = wb2.GetStylesSource();
            ClassicAssert.IsNotNull(ss);

            ClassicAssert.AreEqual(10, st.NumDataFormats);
            ClassicAssert.AreEqual(2, st.Fonts.Count);
            ClassicAssert.AreEqual(2, st.GetFills().Count);
            ClassicAssert.AreEqual(1, st.GetBorders().Count);

            wb2.Close();
        }
        [Test]
        public void IncrementSheetId()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                int sheetId = (int)(wb.CreateSheet() as XSSFSheet).sheet.sheetId;
                ClassicAssert.AreEqual(1, sheetId);
                sheetId = (int)(wb.CreateSheet() as XSSFSheet).sheet.sheetId;
                ClassicAssert.AreEqual(2, sheetId);

                //test file with gaps in the sheetId sequence
                XSSFWorkbook wbBack = XSSFTestDataSamples.OpenSampleWorkbook("47089.xlsm");
                try
                {
                    int lastSheetId = (int)(wbBack.GetSheetAt(wbBack.NumberOfSheets - 1) as XSSFSheet).sheet.sheetId;
                    sheetId = (int)(wbBack.CreateSheet() as XSSFSheet).sheet.sheetId;
                    ClassicAssert.AreEqual(lastSheetId + 1, sheetId);
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
                ClassicAssert.IsNotNull(props);
                //the Application property must be set for new workbooks, see Bugzilla #47559
                ClassicAssert.AreEqual("NPOI", props.ExtendedProperties.GetUnderlyingProperties().Application);

                PackagePropertiesPart opcProps = props.CoreProperties.GetUnderlyingProperties();
                ClassicAssert.IsNotNull(opcProps);

                opcProps.SetTitleProperty("Testing Bugzilla #47460");
                ClassicAssert.AreEqual("NPOI", opcProps.GetCreatorProperty());
                opcProps.SetCreatorProperty("poi-dev@poi.apache.org");

                XSSFWorkbook wbBack = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
                ClassicAssert.AreEqual("NPOI", wbBack.GetProperties().ExtendedProperties.GetUnderlyingProperties().Application);
                opcProps = wbBack.GetProperties().CoreProperties.GetUnderlyingProperties();
                ClassicAssert.AreEqual("Testing Bugzilla #47460", opcProps.GetTitleProperty());
                ClassicAssert.AreEqual("poi-dev@poi.apache.org", opcProps.GetCreatorProperty());
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
        [Test]
        public void Bug47668()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("47668.xlsx");
            IList allPictures = workbook.GetAllPictures();
            ClassicAssert.AreEqual(1, allPictures.Count);

            PackagePartName imagePartName = PackagingUriHelper
                    .CreatePartName("/xl/media/image1.jpeg");
            PackagePart imagePart = workbook.Package.GetPart(imagePartName);
            ClassicAssert.IsNotNull(imagePart);

            foreach (XSSFPictureData pictureData in allPictures)
            {
                PackagePart picturePart = pictureData.GetPackagePart();
                ClassicAssert.AreSame(imagePart, picturePart);
            }

            XSSFSheet sheet0 = (XSSFSheet)workbook.GetSheetAt(0);
            XSSFDrawing Drawing0 = (XSSFDrawing)sheet0.CreateDrawingPatriarch();
            XSSFPictureData pictureData0 = (XSSFPictureData)Drawing0.GetRelations()[0];
            byte[] data0 = pictureData0.Data;
            CRC32 crc0 = new CRC32();
            crc0.Update(data0);

            XSSFSheet sheet1 = workbook.GetSheetAt(1) as XSSFSheet;
            XSSFDrawing Drawing1 = sheet1.CreateDrawingPatriarch() as XSSFDrawing;
            XSSFPictureData pictureData1 = (XSSFPictureData)Drawing1.GetRelations()[0];
            byte[] data1 = pictureData1.Data;
            CRC32 crc1 = new CRC32();
            crc1.Update(data1);

            ClassicAssert.AreEqual(crc0.Value, crc1.Value);
            workbook.Close();
        }

        /**
         * When deleting a sheet make sure that we adjust sheet indices of named ranges
         */
        [Test]
        public void Bug47737()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("47737.xlsx");
            ClassicAssert.AreEqual(2, wb.NumberOfNames);
            ClassicAssert.IsNotNull(wb.GetCalculationChain());

            XSSFName nm0 = (XSSFName)wb.GetAllNames().First();
            ClassicAssert.IsTrue(nm0.GetCTName().IsSetLocalSheetId());
            ClassicAssert.AreEqual(0u, nm0.GetCTName().localSheetId);

            XSSFName nm1 = (XSSFName)wb.GetAllNames().ToArray()[1];
            ClassicAssert.IsTrue(nm1.GetCTName().IsSetLocalSheetId());
            ClassicAssert.AreEqual(1u, nm1.GetCTName().localSheetId);

            wb.RemoveSheetAt(0);
            ClassicAssert.AreEqual(1, wb.NumberOfNames);
            XSSFName nm2 = (XSSFName)wb.GetAllNames().First();
            ClassicAssert.IsTrue(nm2.GetCTName().IsSetLocalSheetId());
            ClassicAssert.AreEqual(0u, nm2.GetCTName().localSheetId);
            //calculation chain is Removed as well
            ClassicAssert.IsNull(wb.GetCalculationChain());

            wb.Close();
        }

        /**
         * Problems with XSSFWorkbook.RemoveSheetAt when workbook Contains chart
         */
        [Test]
        public void Bug47813()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("47813.xlsx");
            ClassicAssert.AreEqual(3, wb1.NumberOfSheets);
            ClassicAssert.IsNotNull(wb1.GetCalculationChain());

            ClassicAssert.AreEqual("Numbers", wb1.GetSheetName(0));
            //the second sheet is of type 'chartsheet'
            ClassicAssert.AreEqual("Chart", wb1.GetSheetName(1));
            ClassicAssert.IsTrue(wb1.GetSheetAt(1) is XSSFChartSheet);
            ClassicAssert.AreEqual("SomeJunk", wb1.GetSheetName(2));

            wb1.RemoveSheetAt(2);
            ClassicAssert.AreEqual(2, wb1.NumberOfSheets);
            ClassicAssert.IsNull(wb1.GetCalculationChain());

            XSSFWorkbook wb2 = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            ClassicAssert.AreEqual(2, wb2.NumberOfSheets);
            ClassicAssert.IsNull(wb2.GetCalculationChain());

            ClassicAssert.AreEqual("Numbers", wb2.GetSheetName(0));
            ClassicAssert.AreEqual("Chart", wb2.GetSheetName(1));

            wb2.Close();
            wb1.Close();
        }

        /**
         * Problems with the count of the number of styles
         *  coming out wrong
         */
        [Test]
        public void Bug49702()
        {
            // First try with a new file
            XSSFWorkbook wb1 = new XSSFWorkbook();

            // Should have one style
            ClassicAssert.AreEqual(1, wb1.NumCellStyles);
            wb1.GetCellStyleAt((short)0);
            ClassicAssert.IsNull(wb1.GetCellStyleAt((short)1),"Shouldn't be able to get style at 0 that doesn't exist");

            // Add another one
            ICellStyle cs = wb1.CreateCellStyle();
            cs.DataFormat = ((short)11);

            // Re-check
            ClassicAssert.AreEqual(2, wb1.NumCellStyles);
            wb1.GetCellStyleAt((short)0);
            wb1.GetCellStyleAt((short)1);
            ClassicAssert.IsNull(wb1.GetCellStyleAt((short)2), "Shouldn't be able to get style at 2 that doesn't exist");

            // Save and reload
            XSSFWorkbook nwb = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            ClassicAssert.AreEqual(2, nwb.NumCellStyles);
            nwb.GetCellStyleAt((short)0);
            nwb.GetCellStyleAt((short)1);
            ClassicAssert.IsNull(nwb.GetCellStyleAt((short)2), "Shouldn't be able to Get style at 2 that doesn't exist");

            // Now with an existing file
            XSSFWorkbook wb2 = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            ClassicAssert.AreEqual(3, wb2.NumCellStyles);
            wb2.GetCellStyleAt((short)0);
            wb2.GetCellStyleAt((short)1);
            wb2.GetCellStyleAt((short)2);
            ClassicAssert.IsNull(nwb.GetCellStyleAt((short)3), "Shouldn't be able to Get style at 3 that doesn't exist");

            wb2.Close();
            wb1.Close();            nwb.Close();
        }
        [Test]
        public void RecalcId()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ClassicAssert.IsFalse(wb.GetForceFormulaRecalculation());
            CT_Workbook ctWorkbook = wb.GetCTWorkbook();
            ClassicAssert.IsFalse(ctWorkbook.IsSetCalcPr());

            wb.SetForceFormulaRecalculation(true); // resets the EngineId flag to zero

            CT_CalcPr calcPr = ctWorkbook.calcPr;
            ClassicAssert.IsNotNull(calcPr);
            ClassicAssert.AreEqual(0, (int)calcPr.calcId);

            calcPr.calcId = 100;
            ClassicAssert.IsTrue(wb.GetForceFormulaRecalculation());

            wb.SetForceFormulaRecalculation(true); // resets the EngineId flag to zero
            ClassicAssert.AreEqual(0, (int)calcPr.calcId);
            ClassicAssert.IsFalse(wb.GetForceFormulaRecalculation());

            // calcMode="manual" is unset when forceFormulaRecalculation=true
            calcPr.calcMode = (ST_CalcMode.manual);
            wb.SetForceFormulaRecalculation(true);
            ClassicAssert.AreEqual(ST_CalcMode.auto, calcPr.calcMode);
        }
        [Test]
        public void ChangeSheetNameWithSharedFormulas()
        {
            ChangeSheetNameWithSharedFormulas("shared_formulas.xlsx");
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
            workbook.Close();
        }

        private void accessWorkbook(XSSFWorkbook workbook)
        {
            workbook.GetSheetAt(1).SetColumnGroupCollapsed(4, true);
            workbook.GetSheetAt(1).SetColumnGroupCollapsed(4, false);

            ClassicAssert.AreEqual("hello world", workbook.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
            ClassicAssert.AreEqual(2158.08, workbook.GetSheetAt(0).GetColumnWidth(0)); // <-works
        }

        [Test]
        public void Bug48495()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("48495.xlsx");

            assertSheetOrder(wb, "Sheet1");

            ISheet sheet = wb.GetSheetAt(0);
            sheet.ShiftRows(2, sheet.LastRowNum, 1, true, false);
            IRow newRow = sheet.GetRow(2);
            if (newRow == null)
                newRow = sheet.CreateRow(2);
            newRow.CreateCell(0).SetCellValue(" Another Header");
            wb.CloneSheet(0);

            assertSheetOrder(wb, "Sheet1", "Sheet1 (2)");


            IWorkbook read = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            ClassicAssert.IsNotNull(read);
            assertSheetOrder(read, "Sheet1", "Sheet1 (2)");

            read.Close();
            wb.Close();
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

            read.Close();
            workbook.Close();
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

            read.Close();
            workbook.Close();
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

            read.Close();
            workbook.Close();
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

            read.Close();
            workbook.Close();
        }

        [Test]
        public void Bug51158()
        {
            // create a workbook
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFSheet sheet = wb1.CreateSheet("Test Sheet") as XSSFSheet;
            XSSFRow row = sheet.CreateRow(2) as XSSFRow;
            XSSFCell cell = row.CreateCell(3) as XSSFCell;
            cell.SetCellValue("test1");

            //XSSFCreationHelper helper = workbook.GetCreationHelper();
            //cell.Hyperlink=(/*setter*/helper.CreateHyperlink(0));

            XSSFComment comment = (sheet.CreateDrawingPatriarch() as XSSFDrawing).CreateCellComment(new XSSFClientAnchor()) as XSSFComment;
            ClassicAssert.IsNotNull(comment);
            comment.SetString("some comment");

            //        ICellStyle cs = workbook.CreateCellStyle();
            //        cs.ShrinkToFit=(/*setter*/false);
            //        row.CreateCell(0).CellStyle=(/*setter*/cs);

            // write the first excel file
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            ClassicAssert.IsNotNull(wb2);
            sheet = wb2.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(2) as XSSFRow;
            ClassicAssert.AreEqual("test1", row.GetCell(3).StringCellValue);
            ClassicAssert.IsNull(row.GetCell(4));

            // add a new cell to the sheet
            cell = row.CreateCell(4) as XSSFCell;
            cell.SetCellValue("test2");

            // write the second excel file
            XSSFWorkbook wb3 = XSSFTestDataSamples.WriteOutAndReadBack(wb2) as XSSFWorkbook;
            ClassicAssert.IsNotNull(wb3);
            sheet = wb3.GetSheetAt(0) as XSSFSheet;
            row = sheet.GetRow(2) as XSSFRow;

            ClassicAssert.AreEqual("test1", row.GetCell(3).StringCellValue);
            ClassicAssert.AreEqual("test2", row.GetCell(4).StringCellValue);

            wb3.Close();
            wb2.Close();
            wb1.Close();
        }

        [Test]
        public void Bug51158a()
        {
            // create a workbook
            XSSFWorkbook workbook = new XSSFWorkbook();
            try
            {
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

                ClassicAssert.AreEqual(1, countMatches(str, "<worksheet"));
            }
            finally
            {
                workbook.Close();
            }

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
                ClassicAssert.IsTrue(ctWb.IsSetPivotCaches());
                ClassicAssert.AreSame(pivotCache, ctWb.pivotCaches.GetPivotCacheArray(0));
                ClassicAssert.AreEqual("0", pivotCache.id);
            }
            finally
            {
                wb.Close();
            }
        }

        protected void SetPivotData(XSSFWorkbook wb)
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

            AreaReference source = wb.GetCreationHelper().CreateAreaReference("A1:B2");
            sheet.CreatePivotTable(source, new CellReference("H5"));
        }

        [Test]
        public void LoadWorkbookWithPivotTable()
        {
            String fileName = Path.Combine(TestContext.CurrentContext.TestDirectory, "ooxml-pivottable.xlsx");

            XSSFWorkbook wb = new XSSFWorkbook();
            SetPivotData(wb);

            FileStream fileOut = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
            wb.Write(fileOut);
            fileOut.Close();

            XSSFWorkbook wb2 = (XSSFWorkbook)WorkbookFactory.Create(fileName);
            ClassicAssert.IsTrue(wb2.PivotTables.Count == 1);
        }

        [Test]
        public void AddPivotTableToWorkbookWithLoadedPivotTable()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            String fileName = "ooxml-pivottable.xlsx";

            XSSFWorkbook wb = new XSSFWorkbook();
            SetPivotData(wb);

            FileStream fileOut = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
            wb.Write(fileOut);
            fileOut.Close();

            XSSFWorkbook wb2 = (XSSFWorkbook)WorkbookFactory.Create(fileName);
            SetPivotData(wb2);
            ClassicAssert.IsTrue(wb2.PivotTables.Count == 2);
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

                ClassicAssert.AreEqual(idx2, wb.FirstVisibleTab);
                ClassicAssert.AreEqual(idx3, wb.ActiveSheetIndex);

                IWorkbook wbBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);

                sheet2 = wbBack.GetSheetAt(idx2);
                sheet3 = wbBack.GetSheetAt(idx3);
                ClassicAssert.AreEqual(idx2, wb.FirstVisibleTab);
                ClassicAssert.AreEqual(idx3, wb.ActiveSheetIndex);

                wbBack.Close();
            }
            finally
            {
                wb.Close();
            }
        }

        /**
         * Tests that we can save a workbook with macros and reload it.
         */
        [Test]
        public void TestSetVBAProject()
        {
            FileInfo file;
            byte[] allBytes = new byte[256];
            for (int i = 0; i < 256; i++)
            {
                allBytes[i] = (byte)(i - 128);
            }

            using(XSSFWorkbook wb1 = new XSSFWorkbook())
            {
                wb1.CreateSheet();
                wb1.SetVBAProject(new ByteArrayInputStream(allBytes));
                file = TempFile.CreateTempFile("poi-", ".xlsm");
                using(Stream out1 = new FileStream(file.FullName, FileMode.Open, FileAccess.ReadWrite))
                {
                    wb1.Write(out1);
                }
            }

            if(file != null)
                file.Refresh();

            // Check the package contains what we'd expect it to
            OPCPackage pkg = OPCPackage.Open(file);
            PackagePart wbPart = pkg.GetPart(PackagingUriHelper.CreatePartName("/xl/workbook.xml"));
            ClassicAssert.IsTrue(wbPart.HasRelationships);
            PackageRelationshipCollection relationships = wbPart.Relationships.GetRelationships(XSSFRelation.VBA_MACROS.Relation);
            ClassicAssert.AreEqual(1, relationships.Size);
            ClassicAssert.AreEqual(XSSFRelation.VBA_MACROS.DefaultFileName, relationships.GetRelationship(0).TargetUri.ToString());
            PackagePart vbaPart = pkg.GetPart(PackagingUriHelper.CreatePartName(XSSFRelation.VBA_MACROS.DefaultFileName));
            ClassicAssert.IsNotNull(vbaPart);
            ClassicAssert.IsFalse(vbaPart.IsRelationshipPart);
            ClassicAssert.AreEqual(XSSFRelation.VBA_MACROS.ContentType, vbaPart.ContentType);
            byte[] fromFile = IOUtils.ToByteArray(vbaPart.GetInputStream());
            CollectionAssert.AreEqual(allBytes, fromFile);
            // Load back the XSSFWorkbook just to check nothing explodes
            XSSFWorkbook wb2 = new XSSFWorkbook(pkg);
            ClassicAssert.AreEqual(1, wb2.NumberOfSheets);
            ClassicAssert.AreEqual(XSSFWorkbookType.XLSM, wb2.WorkbookType);
            pkg.Close();
        }


        [Test]
        public void TestBug54399()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("54399.xlsx");

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                //System.out.println("i:" + i);
                workbook.SetSheetName(i, "SheetRenamed" + (i + 1));
            }

            workbook.Close();
        }


        /**
         *  Iterator<XSSFSheet> XSSFWorkbook.iterator was committed in r700472 on 2008-09-30
         *  and has been replaced with Iterator<Sheet> XSSFWorkbook.iterator
         * 
         *  In order to make code for looping over sheets in workbooks standard, regardless
         *  of the type of workbook (HSSFWorkbook, XSSFWorkbook, SXSSFWorkbook), the previously
         *  available Iterator<XSSFSheet> iterator and Iterator<XSSFSheet> sheetIterator
         *  have been replaced with Iterator<Sheet>  {@link #iterator} and
         *  Iterator<Sheet> {@link #sheetIterator}. This makes iterating over sheets in a workbook
         *  similar to iterating over rows in a sheet and cells in a row.
         *  
         *  Note: this breaks backwards compatibility! Existing codebases will need to
         *  upgrade their code with either of the following options presented in this test case.
         *  
         */
        [Test]
        public void Bug58245_XSSFSheetIterator()
        {
            XSSFWorkbook wb = new XSSFWorkbook();

            wb.CreateSheet();

            // =====================================================================
            // Case 1: Existing code uses XSSFSheet for-each loop
            // =====================================================================
            // Original code (no longer valid)
            /*
            for (XSSFSheet sh : wb) {
                sh.createRow(0);
            }
            */

            // Option A:
            foreach (XSSFSheet sh in wb)
            {
                sh.CreateRow(0);
            }

            // Option B (preferred for new code):
            foreach (ISheet sh in wb)
            {
                sh.CreateRow(0);
            }

            // =====================================================================
            // Case 2: Existing code creates an iterator variable
            // =====================================================================
            // Original code (no longer valid)
            /*
            Iterator<XSSFSheet> it = wb.iterator();
            XSSFSheet sh = it.next();
            sh.createRow(0);
            */

            // Option A:
            {
                IEnumerator<ISheet> it = wb.GetEnumerator();
                it.MoveNext();
                XSSFSheet sh = it.Current as XSSFSheet;
                sh.CreateRow(0);
            }

            // Option B:
            {
                //IEnumerator<XSSFSheet> it = wb.XssfSheetIterator();
                //XSSFSheet sh = it.Current;
                //sh.CreateRow(0);
            }

            // Option C (preferred for new code):
            {
                IEnumerator<ISheet> it = wb.GetEnumerator();
                it.MoveNext();
                ISheet sh = it.Current;
                sh.CreateRow(0);
            }
            wb.Close();
        }

        [Ignore("This unit test may fail on Windows")]
        public void TestBug56957CloseWorkbook()
        {
            FileInfo file = TempFile.CreateTempFile("TestBug56957_", ".xlsx");
            //String dateExp = "Sun Nov 09 00:00:00 CET 2014";
            DateTime dateExp = LocaleUtil.GetLocaleCalendar(2014, 11, 9);
            try
            {
                // as the file is written to, we make a copy before actually working on it
                FileHelper.CopyFile(HSSFTestDataSamples.GetSampleFile("56957.xlsx"), file);

                ClassicAssert.IsTrue(file.Exists);

                // read-only mode works!
                using(var workbook = WorkbookFactory.Create(OPCPackage.Open(file, PackageAccess.READ)))
                {
                    var dateAct = workbook.GetSheetAt(0).GetRow(0).GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).DateCellValue;
                    ClassicAssert.AreEqual(dateExp, dateAct);
                }

                using(var workbook = WorkbookFactory.Create(OPCPackage.Open(file, PackageAccess.READ)))
                {
                    var dateAct = workbook.GetSheetAt(0).GetRow(0).GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).DateCellValue;
                    ClassicAssert.AreEqual(dateExp, dateAct);
                }

                // now check read/write mode
                using(var workbook = WorkbookFactory.Create(OPCPackage.Open(file, PackageAccess.READ_WRITE)))
                {
                    var dateAct = workbook.GetSheetAt(0).GetRow(0).GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).DateCellValue;
                    ClassicAssert.AreEqual(dateExp, dateAct);
                }

                using(var workbook = WorkbookFactory.Create(OPCPackage.Open(file, PackageAccess.READ_WRITE)))
                {
                    var dateAct = workbook.GetSheetAt(0).GetRow(0).GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).DateCellValue;
                    ClassicAssert.AreEqual(dateExp, dateAct);
                }
            }
            finally
            {
                ClassicAssert.IsTrue(file.Exists);

                file.Delete();
                file.Refresh();

                ClassicAssert.IsFalse(file.Exists);
            }
        }

        [Test]
        public void CloseDoesNotModifyWorkbook()
        {
            String filename = "SampleSS.xlsx";
            FileInfo file = POIDataSamples.GetSpreadSheetInstance().GetFileInfo(filename);
            IWorkbook wb;

            // Some tests commented out because close() modifies the file
            // See bug 58779

            // String
            //wb = new XSSFWorkbook(file.Path);
            //assertCloseDoesNotModifyFile(filename, wb);

            // File
            //wb = new XSSFWorkbook(file);
            //assertCloseDoesNotModifyFile(filename, wb);

            // InputStream
            wb = new XSSFWorkbook(file.OpenRead());
            assertCloseDoesNotModifyFile(filename, wb);

            // OPCPackage
            //wb = new XSSFWorkbook(OPCPackage.open(file));
            //assertCloseDoesNotModifyFile(filename, wb);
        }
        [Test]
        public void TestCloseBeforeWrite()
        {
            IWorkbook wb = new XSSFWorkbook();
            wb.CreateSheet("somesheet");
            // test what happens if we close the Workbook before we write it out
            wb.Close();
            try
            {
                XSSFTestDataSamples.WriteOutAndReadBack(wb);
                Assert.Fail("Expecting IOException here");
            }
            catch (RuntimeException e)
            {
                // expected here
                ClassicAssert.IsTrue(e.InnerException is IOException, "Had: " + e.InnerException);
            }
        }

        /**
         * See bug #57840 test data tables
         */
        [Test]
        public void GetTable()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithTable.xlsx");
            XSSFTable table1 = wb.GetTable("Tabella1");
            ClassicAssert.IsNotNull(table1, "Tabella1 was not found in workbook");
            ClassicAssert.AreEqual("Tabella1", table1.Name, "Table name");
            ClassicAssert.AreEqual("Foglio1", table1.SheetName, "Sheet name");
            // Table lookup should be case-insensitive
            ClassicAssert.AreSame(table1, wb.GetTable("TABELLA1"), "Case insensitive table name lookup");
            // If workbook does not contain any data tables matching the provided name, getTable should return null
            ClassicAssert.IsNull(wb.GetTable(null), "Null table name should not throw NPE");
            ClassicAssert.IsNull(wb.GetTable("Foglio1"), "Should not be able to find non-existent table");
            // If a table is added after getTable is called it should still be reachable by XSSFWorkbook.getTable
            // This test makes sure that if any caching is done that getTable never uses a stale cache
            XSSFTable table2 = (wb.GetSheet("Foglio2") as XSSFSheet).CreateTable();
            table2.Name = "Table2";
            ClassicAssert.AreSame(table2, wb.GetTable("Table2"), "Did not find Table2");

            // If table name is modified after getTable is called, the table can only be found by its new name
            // This test makes sure that if any caching is done that getTable never uses a stale cache
            table1.Name = "Table1";
            ClassicAssert.AreSame(table1, wb.GetTable("TABLE1"), "Did not find Tabella1 renamed to Table1");
            wb.Close();
        }

        [Test]
        public void TestRemoveSheet()
        {
            // Test removing a sheet maintains the named ranges correctly
            XSSFWorkbook wb = new XSSFWorkbook();
            wb.CreateSheet("Sheet1");
            wb.CreateSheet("Sheet2");
            XSSFName sheet1Name = wb.CreateName() as XSSFName;
            sheet1Name.NameName = "name1";
            sheet1Name.SheetIndex = 0;
            sheet1Name.RefersToFormula = "Sheet1!$A$1";
            XSSFName sheet2Name = wb.CreateName() as XSSFName;
            sheet2Name.NameName = "name1";
            sheet2Name.SheetIndex = 1;
            sheet2Name.RefersToFormula = "Sheet2!$A$1";
            ClassicAssert.IsTrue(wb.GetAllNames().Contains(sheet1Name));
            ClassicAssert.IsTrue(wb.GetAllNames().Contains(sheet2Name));
            ClassicAssert.AreEqual(2, wb.GetNames("name1").Count);
            ClassicAssert.AreEqual(sheet1Name, wb.GetNames("name1")[0]);
            ClassicAssert.AreEqual(sheet2Name, wb.GetNames("name1")[1]);
            // Remove sheet1, we should only have sheet2Name now
            wb.RemoveSheetAt(0);
            ClassicAssert.IsFalse(wb.GetAllNames().Contains(sheet1Name));
            ClassicAssert.IsTrue(wb.GetAllNames().Contains(sheet2Name));
            ClassicAssert.AreEqual(1, wb.GetNames("name1").Count);
            ClassicAssert.AreEqual(sheet2Name, wb.GetNames("name1")[0]);
            // Check by index as well for sanity
            ClassicAssert.AreEqual(1, wb.NumberOfNames);
            ClassicAssert.AreEqual(sheet2Name, wb.GetName(sheet2Name.NameName));
            wb.Close();
        }

        [Test]
        public void TestRemoveSheetMethod()
        {
            using (XSSFWorkbook wb = new XSSFWorkbook())
            {
                var sheet1 = wb.CreateSheet("Sheet1");
                var sheet2 = wb.CreateSheet("Sheet2");

                ClassicAssert.True(wb.Remove(sheet2));
                ClassicAssert.AreEqual(1, wb.NumberOfSheets);
                ClassicAssert.AreEqual("Sheet1", wb.GetSheetName(0));
                ClassicAssert.AreEqual(sheet1, wb.GetSheet("Sheet1"));

                using (var wbCopy = XSSFTestDataSamples.WriteOutAndReadBack(wb))
                {
                    ClassicAssert.AreEqual(1, wbCopy.NumberOfSheets);
                    ClassicAssert.AreEqual("Sheet1", wb.GetSheetName(0));
                }
            }
        }


        /// <summary>
        /// Issue 1252 https://github.com/nissl-lab/npoi/issues/1252
        /// </summary>
        [Test]
        public void TestSavingXlsxTwiceWithBmpPictures()
        {
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("Issue1252.xlsx");
            byte[] bmpData = GenerateBitmap(100, 100);
            InsertPicture(wb, bmpData, PictureType.BMP);

            using XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            InsertPicture(wb2, bmpData, PictureType.BMP);

            using XSSFWorkbook wb3 = XSSFTestDataSamples.WriteOutAndReadBack(wb2);
            InsertPicture(wb3, bmpData, PictureType.BMP);
            XSSFTestDataSamples.WriteOutAndClose(wb3);
        }

        private static void InsertPicture(IWorkbook wb, byte[] data, PictureType picType)
        {
            ISheet sheet = wb.GetSheetAt(0);
            IDrawing<IShape> patriarch = sheet.DrawingPatriarch;
            XSSFClientAnchor anchor = new(500, 200, 0, 0, 2, 2, 4, 7) {
                AnchorType = AnchorType.MoveDontResize
            };
            int imageId = wb.AddPicture(data, picType);
            XSSFPicture picture = (XSSFPicture)patriarch.CreatePicture(anchor, imageId);
            picture.LineStyle = LineStyle.DashDotGel;
            picture.Resize();
        }

        private static byte[] GenerateBitmap(int width, int height) {
            // BMP file header
            byte[] header = new byte[]
        {
            0x42, 0x4D, // BM (Bitmap identifier)
            0x36, 0x00, 0x0C, 0x00, // File size (54 + width * height)
            0x00, 0x00, // Reserved
            0x00, 0x00, // Reserved
            0x36, 0x00, 0x00, 0x00, // Offset to pixel array
            0x28, 0x00, 0x00, 0x00, // Header size (40 bytes)
            0x64, 0x00, 0x00, 0x00, // Width
            0x64, 0x00, 0x00, 0x00, // Height
            0x01, 0x00, // Planes
            0x18, 0x00, // Bits per pixel (24-bit)
            0x00, 0x00, 0x00, 0x00, // Compression (none)
            0x00, 0x00, 0x00, 0x00, // Image size (can be 0 for uncompressed images)
            0x00, 0x00, 0x00, 0x00, // X pixels per meter
            0x00, 0x00, 0x00, 0x00, // Y pixels per meter
            0x00, 0x00, 0x00, 0x00, // Colors in color table
            0x00, 0x00, 0x00, 0x00, // Important color count
        };

            // Set width and height in header
            BitConverter.GetBytes(width).CopyTo(header, 0x12);
            BitConverter.GetBytes(height).CopyTo(header, 0x16);

            // Create a byte array to hold the pixel data
            int pixelDataSize = width * height * 3; // 3 bytes per pixel for 24-bit color
            byte[] pixels = new byte[pixelDataSize];

            // Set all pixels to black
            for(int i = 0; i < pixelDataSize; i += 3)
            {
                pixels[i] = 0x00; // Blue
                pixels[i + 1] = 0x00; // Green
                pixels[i + 2] = 0x00; // Red
            }

            // Concatenate header and pixel data
            byte[] bmpData = new byte[header.Length + pixels.Length];
            header.CopyTo(bmpData, 0);
            pixels.CopyTo(bmpData, header.Length);

            return bmpData;
        }

        [Test]
        public async Task TestBasicWriteAsync()
        {
            // Create a simple workbook with data
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("TestSheet");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("Async Test Data");

            // Test async write to memory stream
            using (MemoryStream asyncStream = new MemoryStream())
            {
                await wb.WriteAsync(asyncStream);
                ClassicAssert.Greater(asyncStream.Length, 0, "Async write should produce data");
                
                // Verify the written data can be read back
                asyncStream.Position = 0;
                using (XSSFWorkbook readWb = new XSSFWorkbook(asyncStream))
                {
                    ISheet readSheet = readWb.GetSheetAt(0);
                    ClassicAssert.AreEqual("TestSheet", readSheet.SheetName);
                    ClassicAssert.AreEqual("Async Test Data", readSheet.GetRow(0).GetCell(0).StringCellValue);
                }
            }
            
            wb.Dispose();
        }

        [Test]
        public async Task TestWriteAsyncVsSyncComparison()
        {
            // Create identical workbooks
            XSSFWorkbook wb1 = new XSSFWorkbook();
            XSSFWorkbook wb2 = new XSSFWorkbook();
            
            // Add identical data to both
            for (int i = 0; i < 2; i++)
            {
                var workbooks = new[] { wb1, wb2 };
                foreach (var wb in workbooks)
                {
                    ISheet sheet = wb.CreateSheet($"Sheet{i}");
                    for (int row = 0; row < 5; row++)
                    {
                        IRow r = sheet.CreateRow(row);
                        for (int col = 0; col < 3; col++)
                        {
                            ICell cell = r.CreateCell(col);
                            cell.SetCellValue($"Data_{row}_{col}");
                        }
                    }
                }
            }

            // Write using sync method
            byte[] syncData;
            using (MemoryStream syncStream = new MemoryStream())
            {
                wb1.Write(syncStream);
                syncData = syncStream.ToArray();
            }

            // Write using async method
            byte[] asyncData;
            using (MemoryStream asyncStream = new MemoryStream())
            {
                await wb2.WriteAsync(asyncStream);
                asyncData = asyncStream.ToArray();
            }

            // Compare results - they should be identical
            ClassicAssert.AreEqual(syncData.Length, asyncData.Length, "Sync and async should produce same size output");
            
            // Verify both can be read successfully
            using (var syncWb = new XSSFWorkbook(new MemoryStream(syncData)))
            using (var asyncWb = new XSSFWorkbook(new MemoryStream(asyncData)))
            {
                ClassicAssert.AreEqual(syncWb.NumberOfSheets, asyncWb.NumberOfSheets);
                ClassicAssert.AreEqual("Data_2_1", syncWb.GetSheetAt(0).GetRow(2).GetCell(1).StringCellValue);
                ClassicAssert.AreEqual("Data_2_1", asyncWb.GetSheetAt(0).GetRow(2).GetCell(1).StringCellValue);
            }
            
            wb1.Dispose();
            wb2.Dispose();
        }

        [Test]
        public async Task TestWriteAsyncWithCancellation()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("TestSheet");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("Cancellation Test");

            // Test with non-cancelled token
            using (var cts = new CancellationTokenSource())
            using (MemoryStream stream = new MemoryStream())
            {
                await wb.WriteAsync(stream, false, cts.Token);
                ClassicAssert.Greater(stream.Length, 0, "Write should complete successfully with non-cancelled token");
            }

            wb.Dispose();
        }
    }
}
