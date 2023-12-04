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
    using System.IO;
    using System.Collections;

    using TestCases.HSSF;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.SS.Formula;
    using NPOI.Util;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    using NPOI.DDF;
    using TestCases.SS.UserModel;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.Util;
    using System.Collections.Generic;
    using System.Text;
    using NPOI.HSSF;
    using System.Threading;
    using System.Globalization;
    using NPOI.SS;

    /**
*
*/
    [TestFixture]
    public class TestHSSFWorkbook : BaseTestWorkbook
    {
        public TestHSSFWorkbook()
            : base(HSSFITestDataProvider.Instance)
        {
        }
        /**
     * gives test code access to the {@link InternalWorkbook} within {@link HSSFWorkbook}
     */
        public static InternalWorkbook GetInternalWorkbook(HSSFWorkbook wb)
        {
            return wb.Workbook;
        }
        private static HSSFWorkbook OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleWorkbook(sampleFileName);
        }

        /**
         * Tests for {@link HSSFWorkbook#isHidden()} etc
         * @throws IOException 
         */
        [Test]
        public void Hidden()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            WindowOneRecord w1 = wb.Workbook.WindowOne;

            Assert.AreEqual(false, wb.IsHidden);
            Assert.AreEqual(false, w1.Hidden);

            wb.IsHidden = (true);
            Assert.AreEqual(true, wb.IsHidden);
            Assert.AreEqual(true, w1.Hidden);

            HSSFWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            w1 = wbBack.Workbook.WindowOne;

            wbBack.IsHidden = (true);
            Assert.AreEqual(true, wbBack.IsHidden);
            Assert.AreEqual(true, w1.Hidden);

            wbBack.IsHidden = (false);
            Assert.AreEqual(false, wbBack.IsHidden);
            Assert.AreEqual(false, w1.Hidden);

            wbBack.Close();
            wb.Close();
        }

        [Test]
        [Ignore("not found in poi")]
        public void CaseInsensitiveNames()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            ISheet originalSheet = b.CreateSheet("Sheet1");
            ISheet fetchedSheet = b.GetSheet("sheet1");
            if (fetchedSheet == null)
            {
                throw new AssertionException("Identified bug 44892");
            }
            Assert.AreEqual(originalSheet, fetchedSheet);
            try
            {
                b.CreateSheet("sHeeT1");
                Assert.Fail("should have thrown exceptiuon due to duplicate sheet name");
            }
            catch (ArgumentException e)
            {
                // expected during successful Test
                Assert.AreEqual("The workbook already contains a sheet of this name", e.Message);
            }
        }
        [Test]
        [Ignore("not found in poi")]
        public void DuplicateNames()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            b.CreateSheet("Sheet1");
            b.CreateSheet();
            b.CreateSheet("name1");
            try
            {
                b.CreateSheet("name1");
                Assert.Fail();
            }
            catch (ArgumentException)// pass
            {
            }
            b.CreateSheet();
            try
            {
                b.SetSheetName(3, "name1");
                Assert.Fail();
            }
            catch (ArgumentException)// pass
            {
            }

            try
            {
                b.SetSheetName(3, "name1");
                Assert.Fail();
            }
            catch (ArgumentException)// pass
            {
            }

            b.SetSheetName(3, "name2");
            b.SetSheetName(3, "name2");
            b.SetSheetName(3, "name2");

            HSSFWorkbook c = new HSSFWorkbook();
            c.CreateSheet("Sheet1");
            c.CreateSheet("Sheet2");
            c.CreateSheet("Sheet3");
            c.CreateSheet("Sheet4");

        }

        [Test]
        [Ignore("not found in poi")]
        public new void TestSheetSelection()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            b.CreateSheet("Sheet One");
            b.CreateSheet("Sheet Two");
            b.SetActiveSheet(1);
            b.SetSelectedTab(1);
            b.FirstVisibleTab = (1);
            Assert.AreEqual(1, b.ActiveSheetIndex);
            Assert.AreEqual(1, b.FirstVisibleTab);
        }

        [Test]
        public void ReadWriteWithCharts()
        {
            ISheet s;
            // Single chart, two sheets
            HSSFWorkbook b1 = HSSFTestDataSamples.OpenSampleWorkbook("44010-SingleChart.xls");
            Assert.AreEqual(2, b1.NumberOfSheets);
            Assert.AreEqual("Graph2", b1.GetSheetName(1));
            s = b1.GetSheetAt(1);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
            // Has chart on 1st sheet??
            // FIXME
            Assert.IsNotNull(b1.GetSheetAt(0).DrawingPatriarch);
            Assert.IsNull(b1.GetSheetAt(1).DrawingPatriarch);
            Assert.IsFalse((b1.GetSheetAt(0).DrawingPatriarch as HSSFPatriarch).ContainsChart());
            b1.Close();
            // We've now called getDrawingPatriarch() so
            //  everything will be all screwy
            // So, start again
            HSSFWorkbook b2 = HSSFTestDataSamples.OpenSampleWorkbook("44010-SingleChart.xls");
            HSSFWorkbook b3 = HSSFTestDataSamples.WriteOutAndReadBack(b2);
            b2.Close();
            Assert.AreEqual(2, b3.NumberOfSheets);
            s = b3.GetSheetAt(1) as HSSFSheet;
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
            b3.Close();
            // Two charts, three sheets
            HSSFWorkbook b4 = HSSFTestDataSamples.OpenSampleWorkbook("44010-TwoCharts.xls");
            Assert.AreEqual(3, b4.NumberOfSheets);
            s = b4.GetSheetAt(1) as HSSFSheet;
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
            s = b4.GetSheetAt(2) as HSSFSheet;
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
            // Has chart on 1st sheet??
            // FIXME
            Assert.IsNotNull(b4.GetSheetAt(0).DrawingPatriarch);
            Assert.IsNull(b4.GetSheetAt(1).DrawingPatriarch);
            Assert.IsNull(b4.GetSheetAt(2).DrawingPatriarch);
            Assert.IsFalse((b4.GetSheetAt(0).DrawingPatriarch as HSSFPatriarch).ContainsChart());
            b4.Close();
            // We've now called getDrawingPatriarch() so
            //  everything will be all screwy
            // So, start again
            HSSFWorkbook b5 = HSSFTestDataSamples.OpenSampleWorkbook("44010-TwoCharts.xls");
            HSSFWorkbook b6 = HSSFTestDataSamples.WriteOutAndReadBack(b5);
            b5.Close();
            Assert.AreEqual(3, b6.NumberOfSheets);
            s = b6.GetSheetAt(1) as HSSFSheet;
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
            s = b6.GetSheetAt(2) as HSSFSheet;
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
            b6.Close();

        }

        private static HSSFWorkbook WriteRead(HSSFWorkbook b)
        {
            return HSSFTestDataSamples.WriteOutAndReadBack(b);
        }

        [Test]
        public void SelectedSheet_bug44523()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet1 = wb.CreateSheet("Sheet1");
            ISheet sheet2 = wb.CreateSheet("Sheet2");
            ISheet sheet3 = wb.CreateSheet("Sheet3");
            ISheet sheet4 = wb.CreateSheet("Sheet4");

            ConfirmActiveSelected(sheet1, true);
            ConfirmActiveSelected(sheet2, false);
            ConfirmActiveSelected(sheet3, false);
            ConfirmActiveSelected(sheet4, false);

            wb.SetSelectedTab(1);

            // Demonstrate bug 44525:
            // Well... not quite, since isActive + isSelected were also Added in the same bug fix
            if (sheet1.IsSelected)
            {
                throw new AssertionException("Identified bug 44523 a");
            }
            wb.SetActiveSheet(1);
            if (sheet1.IsActive)
            {
                throw new AssertionException("Identified bug 44523 b");
            }

            ConfirmActiveSelected(sheet1, false);
            ConfirmActiveSelected(sheet2, true);
            ConfirmActiveSelected(sheet3, false);
            ConfirmActiveSelected(sheet4, false);
        }
        private static List<int> arrayToList(int[] array)
        {
            List<int> list = new List<int>(array.Length);
            foreach (int element in array)
            {
                list.Add(element);
            }
            return list;
        }

        private static void assertCollectionsEquals(List<int> expected, List<int> actual)
        {
            Assert.AreEqual(expected.Count, actual.Count, "size");
            foreach (int e in expected)
            {
                Assert.IsTrue(actual.Contains(e));
            }
            foreach (int a in actual)
            {
                Assert.IsTrue(expected.Contains(a));
            }
        }
        [Test]
        public void SelectMultiple()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet0 = wb.CreateSheet("Sheet0") as HSSFSheet;
            HSSFSheet sheet1 = wb.CreateSheet("Sheet1") as HSSFSheet;
            HSSFSheet sheet2 = wb.CreateSheet("Sheet2") as HSSFSheet;
            HSSFSheet sheet3 = wb.CreateSheet("Sheet3") as HSSFSheet;
            HSSFSheet sheet4 = wb.CreateSheet("Sheet4") as HSSFSheet;
            HSSFSheet sheet5 = wb.CreateSheet("Sheet5") as HSSFSheet;

            List<int> selected = arrayToList(new int[] { 0, 2, 3 });
            wb.SetSelectedTabs(selected);

            CollectionAssert.AreEqual(selected, wb.GetSelectedTabs());
            Assert.AreEqual(true, sheet0.IsSelected);
            Assert.AreEqual(false, sheet1.IsSelected);
            Assert.AreEqual(true, sheet2.IsSelected);
            Assert.AreEqual(true, sheet3.IsSelected);
            Assert.AreEqual(false, sheet4.IsSelected);
            Assert.AreEqual(false, sheet5.IsSelected);

            selected = arrayToList(new int[] { 1, 3, 5 });
            wb.SetSelectedTabs(selected);

            // previous selection should be cleared
            CollectionAssert.AreEqual(selected, wb.GetSelectedTabs());
            Assert.AreEqual(false, sheet0.IsSelected);
            Assert.AreEqual(true, sheet1.IsSelected);
            Assert.AreEqual(false, sheet2.IsSelected);
            Assert.AreEqual(true, sheet3.IsSelected);
            Assert.AreEqual(false, sheet4.IsSelected);
            Assert.AreEqual(true, sheet5.IsSelected);
            Assert.AreEqual(true, sheet0.IsActive);
            Assert.AreEqual(false, sheet2.IsActive);

            wb.SetActiveSheet(2);
            Assert.AreEqual(false, sheet0.IsActive);
            Assert.AreEqual(true, sheet2.IsActive);


            /*{ // helpful if viewing this workbook in excel:
                sheet0.createRow(0).createCell(0).setCellValue(new HSSFRichTextString("Sheet0"));
                sheet1.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Sheet1"));
                sheet2.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Sheet2"));
                sheet3.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Sheet3"));

                try
                {
                    File fOut = TempFile.CreateTempFile("sheetMultiSelect", ".xls");
                    FileOutputStream os = new FileOutputStream(fOut);
                    wb.Write(os);
                    os.Close();
                }
                catch (IOException e)
                {
                    throw new RuntimeException(e);
                }
            }*/
        }

        [Test]
        public void ActiveSheetAfterDelete_bug40414()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet0 = wb.CreateSheet("Sheet0");
            ISheet sheet1 = wb.CreateSheet("Sheet1");
            ISheet sheet2 = wb.CreateSheet("Sheet2");
            ISheet sheet3 = wb.CreateSheet("Sheet3");
            ISheet sheet4 = wb.CreateSheet("Sheet4");

            // Confirm default activation/selection
            ConfirmActiveSelected(sheet0, true);
            ConfirmActiveSelected(sheet1, false);
            ConfirmActiveSelected(sheet2, false);
            ConfirmActiveSelected(sheet3, false);
            ConfirmActiveSelected(sheet4, false);

            wb.SetActiveSheet(3);
            wb.SetSelectedTab(3);

            ConfirmActiveSelected(sheet0, false);
            ConfirmActiveSelected(sheet1, false);
            ConfirmActiveSelected(sheet2, false);
            ConfirmActiveSelected(sheet3, true);
            ConfirmActiveSelected(sheet4, false);

            wb.RemoveSheetAt(3);
            // after removing the only active/selected sheet, another should be active/selected in its place
            if (!sheet4.IsSelected)
            {
                throw new AssertionException("identified bug 40414 a");
            }
            if (!sheet4.IsActive)
            {
                throw new AssertionException("identified bug 40414 b");
            }

            ConfirmActiveSelected(sheet0, false);
            ConfirmActiveSelected(sheet1, false);
            ConfirmActiveSelected(sheet2, false);
            ConfirmActiveSelected(sheet4, true);

            sheet3 = sheet4; // re-align local vars in this Test case

            // Some more cases of removing sheets

            // Starting with a multiple selection, and different active sheet
            wb.SetSelectedTabs(new int[] { 1, 3, });
            wb.SetActiveSheet(2);
            ConfirmActiveSelected(sheet0, false, false);
            ConfirmActiveSelected(sheet1, false, true);
            ConfirmActiveSelected(sheet2, true, false);
            ConfirmActiveSelected(sheet3, false, true);

            // removing a sheet that is not active, and not the only selected sheet
            wb.RemoveSheetAt(3);
            ConfirmActiveSelected(sheet0, false, false);
            ConfirmActiveSelected(sheet1, false, true);
            ConfirmActiveSelected(sheet2, true, false);

            // removing the only selected sheet
            wb.RemoveSheetAt(1);
            ConfirmActiveSelected(sheet0, false, false);
            ConfirmActiveSelected(sheet2, true, true);

            // The last remaining sheet should always be active+selected
            wb.RemoveSheetAt(1);
            ConfirmActiveSelected(sheet0, true, true);
        }

        private static void ConfirmActiveSelected(ISheet sheet, bool expected)
        {
            ConfirmActiveSelected(sheet, expected, expected);
        }


        private static void ConfirmActiveSelected(ISheet sheet,
                bool expectedActive, bool expectedSelected)
        {
            Assert.AreEqual(expectedActive, sheet.IsActive, "active");
            Assert.AreEqual(expectedSelected, sheet.IsSelected, "selected");
        }

        /**
         * If Sheet.GetSize() returns a different result to Sheet.serialize(), this will cause the BOF
         * records to be written with invalid offset indexes.  Excel does not like this, and such 
         * errors are particularly hard to track down.  This Test ensures that HSSFWorkbook throws
         * a specific exception as soon as the situation is detected. See bugzilla 45066
         */
        [Test]
        public void SheetSerializeSizeMisMatch_bug45066()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            InternalSheet sheet = ((HSSFSheet)wb.CreateSheet("Sheet1")).Sheet;
            IList sheetRecords = sheet.Records;
            // one way (of many) to cause the discrepancy is with a badly behaved record:
            sheetRecords.Add(new BadlyBehavedRecord());
            // There is also much logic inside Sheet that (if buggy) might also cause the discrepancy
            try
            {
                wb.GetBytes();
                throw new AssertionException("Identified bug 45066 a");
            }
            catch (InvalidOperationException e)
            {
                // Expected badly behaved sheet record to cause exception
                Assert.IsTrue(e.Message.StartsWith("Actual serialized sheet size"));
            }
        }

        /**
         * Checks that us and IName play nicely with named ranges
         *  that point to deleted sheets
         */
        [Test]
        public void NamesToDeleteSheets()
        {
            HSSFWorkbook b = OpenSample("30978-deleted.xls");
            Assert.AreEqual(3, b.NumberOfNames);

            // Sheet 2 is deleted
            Assert.AreEqual("Sheet1", b.GetSheetName(0));
            Assert.AreEqual("Sheet3", b.GetSheetName(1));

            Area3DPtg ptg;
            NameRecord nr;
            IName n;

            /* ======= Name pointing to deleted sheet ====== */

            // First at low level
            nr = b.Workbook.GetNameRecord(0);
            Assert.AreEqual("On2", nr.NameText);
            Assert.AreEqual(0, nr.SheetNumber);
            Assert.AreEqual(1, nr.ExternSheetNumber);
            Assert.AreEqual(1, nr.NameDefinition.Length);

            ptg = (Area3DPtg)nr.NameDefinition[0];
            Assert.AreEqual(1, ptg.ExternSheetIndex);
            Assert.AreEqual(0, ptg.FirstColumn);
            Assert.AreEqual(0, ptg.FirstRow);
            Assert.AreEqual(0, ptg.LastColumn);
            Assert.AreEqual(2, ptg.LastRow);

            // Now at high level
            n = b.GetNameAt(0);
            Assert.AreEqual("On2", n.NameName);
            Assert.AreEqual("", n.SheetName);
            Assert.AreEqual("#REF!$A$1:$A$3", n.RefersToFormula);


            /* ======= Name pointing to 1st sheet ====== */

            // First at low level
            nr = b.Workbook.GetNameRecord(1);
            Assert.AreEqual("OnOne", nr.NameText);
            Assert.AreEqual(0, nr.SheetNumber);
            Assert.AreEqual(0, nr.ExternSheetNumber);
            Assert.AreEqual(1, nr.NameDefinition.Length);

            ptg = (Area3DPtg)nr.NameDefinition[0];
            Assert.AreEqual(0, ptg.ExternSheetIndex);
            Assert.AreEqual(0, ptg.FirstColumn);
            Assert.AreEqual(2, ptg.FirstRow);
            Assert.AreEqual(0, ptg.LastColumn);
            Assert.AreEqual(3, ptg.LastRow);

            // Now at high level
            n = b.GetNameAt(1);
            Assert.AreEqual("OnOne", n.NameName);
            Assert.AreEqual("Sheet1", n.SheetName);
            Assert.AreEqual("Sheet1!$A$3:$A$4", n.RefersToFormula);


            /* ======= Name pointing to 3rd sheet ====== */

            // First at low level
            nr = b.Workbook.GetNameRecord(2);
            Assert.AreEqual("OnSheet3", nr.NameText);
            Assert.AreEqual(0, nr.SheetNumber);
            Assert.AreEqual(2, nr.ExternSheetNumber);
            Assert.AreEqual(1, nr.NameDefinition.Length);

            ptg = (Area3DPtg)nr.NameDefinition[0];
            Assert.AreEqual(2, ptg.ExternSheetIndex);
            Assert.AreEqual(0, ptg.FirstColumn);
            Assert.AreEqual(0, ptg.FirstRow);
            Assert.AreEqual(0, ptg.LastColumn);
            Assert.AreEqual(1, ptg.LastRow);

            // Now at high level
            n = b.GetNameAt(2);
            Assert.AreEqual("OnSheet3", n.NameName);
            Assert.AreEqual("Sheet3", n.SheetName);
            Assert.AreEqual("Sheet3!$A$1:$A$2", n.RefersToFormula);

            b.Close();
        }

        /**
         * result returned by getRecordSize() differs from result returned by serialize()
         */
        private class BadlyBehavedRecord : Record
        {
            public BadlyBehavedRecord()
            {
                // 
            }
            public override short Sid
            {
                get
                {
                    return unchecked((short)0x777);
                }
            }
            public override int Serialize(int offset, byte[] data)
            {
                return 4;
            }
            public override int RecordSize
            {
                get
                {
                    return 8;
                }
            }
        }

        /**
         * The sample file provided with bug 45582 seems to have one extra byte after the EOFRecord
         */
        [Test]
        public void ExtraDataAfterEOFRecord()
        {
            try
            {
                HSSFTestDataSamples.OpenSampleWorkbook("ex45582-22397.xls");
            }
            catch (RecordFormatException e)
            {
                if (e.InnerException is NPOI.Util.BufferUnderrunException)
                {
                    throw new AssertionException("Identified bug 45582");
                }
            }
        }

        /**
         * Test to make sure that NameRecord.SheetNumber is interpreted as a
         * 1-based sheet tab index (not a 1-based extern sheet index)
         */
        [Test]
        public void FindBuiltInNameRecord()
        {
            // TestRRaC has multiple (3) built-in name records
            // The second print titles name record has SheetNumber==4
            HSSFWorkbook wb1 = HSSFTestDataSamples.OpenSampleWorkbook("testRRaC.xls");
            NameRecord nr;
            Assert.AreEqual(3, wb1.Workbook.NumNames);
            nr = wb1.Workbook.GetNameRecord(2);
            // TODO - render full row and full column refs properly
            Assert.AreEqual("Sheet2!$A$1:$IV$1", HSSFFormulaParser.ToFormulaString(wb1, nr.NameDefinition)); // 1:1

            try
            {
                wb1.GetSheetAt(3).RepeatingRows = (CellRangeAddress.ValueOf("9:12"));
                wb1.GetSheetAt(3).RepeatingColumns = (CellRangeAddress.ValueOf("E:F"));
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Builtin (7) already exists for sheet (4)"))
                {
                    // there was a problem in the code which locates the existing print titles name record 
                    throw new Exception("Identified bug 45720b");
                }
                throw;
            }
            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb1);
            wb1.Close();

            Assert.AreEqual(3, wb2.Workbook.NumNames);
            nr = wb2.Workbook.GetNameRecord(2);
            Assert.AreEqual("Sheet2!E:F,Sheet2!$A$9:$IV$12", HSSFFormulaParser.ToFormulaString(wb2, nr.NameDefinition)); // E:F,9:12
            wb2.Close();
        }
        /**
     * Test that the storage clsid property is preserved
     */
        [Test]
        public void Bug47920()
        {
            POIFSFileSystem fs1 = new POIFSFileSystem(POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("47920.xls"));
            IWorkbook wb = new HSSFWorkbook(fs1);
            ClassID clsid1 = fs1.Root.StorageClsid;

            MemoryStream out1 = new MemoryStream(4096);
            wb.Write(out1, false);
            byte[] bytes = out1.ToArray();
            POIFSFileSystem fs2 = new POIFSFileSystem(new MemoryStream(bytes));
            ClassID clsid2 = fs2.Root.StorageClsid;

            Assert.IsTrue(clsid1.Equals(clsid2));

            fs2.Close();
            wb.Close();
            fs1.Close();
        }

        /**
     * If we try to open an old (pre-97) workbook, we Get a helpful
     *  Exception give to explain what we've done wrong
     */
        [Test]
        public void HelpfulExceptionOnOldFiles()
        {
            Stream excel4 = POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("testEXCEL_4.xls");
            try
            {
                new HSSFWorkbook(excel4);
                Assert.Fail("Shouldn't be able to load an Excel 4 file");
            }
            catch (OldExcelFormatException e)
            {
                POITestCase.AssertContains(e.Message, "BIFF4");
            }
            excel4.Close();

            Stream excel5 = POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("testEXCEL_5.xls");
            try
            {
                new HSSFWorkbook(excel5);
                Assert.Fail("Shouldn't be able to load an Excel 5 file");
            }
            catch (OldExcelFormatException e)
            {
                POITestCase.AssertContains(e.Message, "BIFF8");
            }
            excel5.Close();

            Stream excel95 = POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("testEXCEL_95.xls");
            try
            {
                new HSSFWorkbook(excel95);
                Assert.Fail("Shouldn't be able to load an Excel 95 file");
            }
            catch (OldExcelFormatException e)
            {
                POITestCase.AssertContains(e.Message, "BIFF5");
            }
            excel95.Close();
        }


        /**
         * Tests that we can work with both {@link POIFSFileSystem}
         *  and {@link NPOIFSFileSystem}
         */
        [Test]
        public void DifferentPOIFS()
        {
            //throw new NotImplementedException("class NPOIFSFileSystem is not implemented");
            // Open the two filesystems
            DirectoryNode[] files = new DirectoryNode[2];
            files[0] = (new POIFSFileSystem(HSSFTestDataSamples.OpenSampleFileStream("Simple.xls"))).Root;
            files[1] = (new NPOIFSFileSystem(HSSFTestDataSamples.OpenSampleFileStream("Simple.xls"))).Root;

            // Open without preserving nodes 
            foreach (DirectoryNode dir in files)
            {
                IWorkbook workbook = new HSSFWorkbook(dir, false);
                ISheet sheet = workbook.GetSheetAt(0);
                ICell cell = sheet.GetRow(0).GetCell(0);
                Assert.AreEqual("replaceMe", cell.RichStringCellValue.String);
            }

            // Now re-check with preserving
            foreach (DirectoryNode dir in files)
            {
                IWorkbook workbook = new HSSFWorkbook(dir, true);
                ISheet sheet = workbook.GetSheetAt(0);
                ICell cell = sheet.GetRow(0).GetCell(0);
                Assert.AreEqual("replaceMe", cell.RichStringCellValue.String);
            }
        }

        [Test]
        public void WordDocEmbeddedInXls()
        {
            //throw new NotImplementedException("class NPOIFSFileSystem is not implemented");
            // Open the two filesystems
            DirectoryNode[] files = new DirectoryNode[2];
            files[0] = (new POIFSFileSystem(HSSFTestDataSamples.OpenSampleFileStream("WithEmbeddedObjects.xls"))).Root;
            files[1] = (new NPOIFSFileSystem(HSSFTestDataSamples.OpenSampleFileStream("WithEmbeddedObjects.xls"))).Root;

            // Check the embedded parts
            foreach (DirectoryNode root in files)
            {
                HSSFWorkbook hw = new HSSFWorkbook(root, true);
                IList<HSSFObjectData> objects = hw.GetAllEmbeddedObjects();
                bool found = false;
                foreach (HSSFObjectData embeddedObject in objects)
                {
                    if (embeddedObject.HasDirectoryEntry())
                    {
                        DirectoryEntry dir = embeddedObject.GetDirectory();
                        if (dir is DirectoryNode)
                        {
                            DirectoryNode dNode = (DirectoryNode)dir;
                            if (HasEntry(dNode, "WordDocument"))
                            {
                                found = true;
                            }
                        }
                    }
                }
                Assert.IsTrue(found);
            }
        }

        /**
         * Checks that we can open a workbook with NPOIFS, and write it out
         *  again (via POIFS) and have it be valid
         * @throws IOException
         */
        [Test]
        public void WriteWorkbookFromNPOIFS()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("WithEmbeddedObjects.xls");
            try
            {
                NPOIFSFileSystem fs = new NPOIFSFileSystem(is1);
                try
                {
                    // Start as NPOIFS
                    HSSFWorkbook wb = new HSSFWorkbook(fs.Root, true);
                    Assert.AreEqual(3, wb.NumberOfSheets);
                    Assert.AreEqual("Root xls", wb.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);

                    // Will switch to POIFS
                    wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
                    Assert.AreEqual(3, wb.NumberOfSheets);
                    Assert.AreEqual("Root xls", wb.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
                }
                finally
                {
                    fs.Close();
                }
            }
            finally
            {
                is1.Close();
            }
        }

        [Test]
        public void CellStylesLimit()
        {
            IWorkbook wb = new HSSFWorkbook();
            int numBuiltInStyles = wb.NumCellStyles;
            int MAX_STYLES = 4030;
            int limit = MAX_STYLES - numBuiltInStyles;
            for (int i = 0; i < limit; i++)
            {
                ICellStyle style = wb.CreateCellStyle();
            }

            Assert.AreEqual(MAX_STYLES, wb.NumCellStyles);
            try
            {
                ICellStyle style = wb.CreateCellStyle();
                Assert.Fail("expected exception");
            }
            catch (InvalidOperationException e)
            {
                Assert.AreEqual("The maximum number of cell styles was exceeded. " +
                        "You can define up to 4000 styles in a .xls workbook", e.Message);
            }
            Assert.AreEqual(MAX_STYLES, wb.NumCellStyles);
        }
        [Test]
        public void SetSheetOrderHSSF()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet s1 = wb.CreateSheet("first sheet");
            ISheet s2 = wb.CreateSheet("other sheet");

            IName name1 = wb.CreateName();
            name1.NameName = (/*setter*/"name1");
            name1.RefersToFormula = (/*setter*/"'first sheet'!D1");

            IName name2 = wb.CreateName();
            name2.NameName = (/*setter*/"name2");
            name2.RefersToFormula = (/*setter*/"'other sheet'!C1");


            IRow s1r1 = s1.CreateRow(2);
            ICell c1 = s1r1.CreateCell(3);
            c1.SetCellValue(30);
            ICell c2 = s1r1.CreateCell(2);
            c2.CellFormula = (/*setter*/"SUM('other sheet'!C1,'first sheet'!C1)");

            IRow s2r1 = s2.CreateRow(0);
            ICell c3 = s2r1.CreateCell(1);
            c3.CellFormula = (/*setter*/"'first sheet'!D3");
            ICell c4 = s2r1.CreateCell(2);
            c4.CellFormula = (/*setter*/"'other sheet'!D3");

            // conditional formatting
            ISheetConditionalFormatting sheetCF = s1.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(
                    ComparisonOperator.Between, "'first sheet'!D1", "'other sheet'!D1");

            IConditionalFormattingRule[] cfRules = { rule1 };

            CellRangeAddress[] regions = { new CellRangeAddress(2, 4, 0, 0), // A3:A5
        };
            sheetCF.AddConditionalFormatting(regions, cfRules);

            wb.SetSheetOrder("other sheet", 0);

            // names
            Assert.AreEqual("'first sheet'!D1", wb.GetName("name1").RefersToFormula);
            Assert.AreEqual("'other sheet'!C1", wb.GetName("name2").RefersToFormula);

            // cells
            Assert.AreEqual("SUM('other sheet'!C1,'first sheet'!C1)", c2.CellFormula);
            Assert.AreEqual("'first sheet'!D3", c3.CellFormula);
            Assert.AreEqual("'other sheet'!D3", c4.CellFormula);

            // conditional formatting
            IConditionalFormatting cf = sheetCF.GetConditionalFormattingAt(0);
            Assert.AreEqual("'first sheet'!D1", cf.GetRule(0).Formula1);
            Assert.AreEqual("'other sheet'!D1", cf.GetRule(0).Formula2);
        }

        private bool HasEntry(DirectoryNode dirNode, String entryName)
        {
            try
            {
                dirNode.GetEntry(entryName);
                return true;
            }
            catch (FileNotFoundException)
            {
                return false;
            }
        }

        [Test]
        public void ClonePictures()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SimpleWithImages.xls");
            InternalWorkbook iwb = ((HSSFWorkbook)wb).Workbook;
            iwb.FindDrawingGroup();

            for (int pictureIndex = 1; pictureIndex <= 4; pictureIndex++)
            {
                EscherBSERecord bse = iwb.GetBSERecord(pictureIndex);
                Assert.AreEqual(1, bse.Ref);
            }

            wb.CloneSheet(0);
            for (int pictureIndex = 1; pictureIndex <= 4; pictureIndex++)
            {
                EscherBSERecord bse = iwb.GetBSERecord(pictureIndex);
                Assert.AreEqual(2, bse.Ref);
            }

            wb.CloneSheet(0);
            for (int pictureIndex = 1; pictureIndex <= 4; pictureIndex++)
            {
                EscherBSERecord bse = iwb.GetBSERecord(pictureIndex);
                Assert.AreEqual(3, bse.Ref);
            }
            wb.Close();
        }

        [Test]
        public void ChangeSheetNameWithSharedFormulas()
        {
            ChangeSheetNameWithSharedFormulas("shared_formulas.xls");
        }
        [Test]
        public void EmptyDirectoryNode()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            try
            {
                new HSSFWorkbook(fs).Close();
            }
            catch (ArgumentException ex)
            {
                Assert.IsTrue(ex.Message.StartsWith("The supplied POIFSFileSystem does not contain a BIFF8"));
            }
            finally
            {
                fs.Close();
            }
        }
        [Test]
        public void SelectedSheetshort()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet1 = (HSSFSheet)wb.CreateSheet("Sheet1");
            HSSFSheet sheet2 = (HSSFSheet)wb.CreateSheet("Sheet2");
            HSSFSheet sheet3 = (HSSFSheet)wb.CreateSheet("Sheet3");
            HSSFSheet sheet4 = (HSSFSheet)wb.CreateSheet("Sheet4");

            ConfirmActiveSelected(sheet1, true);
            ConfirmActiveSelected(sheet2, false);
            ConfirmActiveSelected(sheet3, false);
            ConfirmActiveSelected(sheet4, false);

            wb.SetSelectedTab((short)1);

            // Demonstrate bug 44525:
            // Well... not quite, since isActive + isSelected were also Added in the same bug fix
            if (sheet1.IsSelected)
            {
                //throw new AssertionFailedError("Identified bug 44523 a");
                Assert.Fail("Identified bug 44523 a");
            }
            wb.SetActiveSheet(1);
            if (sheet1.IsActive)
            {
                //throw new AssertionFailedError("Identified bug 44523 b");
                Assert.Fail("Identified bug 44523 b");
            }

            ConfirmActiveSelected(sheet1, false);
            ConfirmActiveSelected(sheet2, true);
            ConfirmActiveSelected(sheet3, false);
            ConfirmActiveSelected(sheet4, false);

            Assert.AreEqual(0, wb.FirstVisibleTab);
            wb.FirstVisibleTab = 2;
            Assert.AreEqual(2, wb.FirstVisibleTab);

            wb.Close();
        }

        [Test]
        public void Names()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            try
            {
                wb.GetNameAt(0);
                Assert.Fail("Fails without any defined names");
            }
            catch (InvalidOperationException e)
            {
                Assert.IsTrue(e.Message.Contains("no defined names"), e.Message);
            }

            HSSFName name = (HSSFName)wb.CreateName();
            Assert.IsNotNull(name);

            Assert.IsNull(wb.GetName("somename"));

            name.NameName = ("myname");
            Assert.IsNotNull(wb.GetName("myname"));

            Assert.AreEqual(0, wb.GetNameIndex(name));
            Assert.AreEqual(0, wb.GetNameIndex("myname"));

            try
            {
                wb.GetNameAt(5);
                Assert.Fail("Fails without any defined names");
            }
            catch (ArgumentOutOfRangeException e)
            {
                Assert.IsTrue(e.Message.Contains("outside the allowable range"), e.Message);
            }

            try
            {
                wb.GetNameAt(-3);
                Assert.Fail("Fails without any defined names");
            }
            catch (ArgumentOutOfRangeException e)
            {
                Assert.IsTrue(e.Message.Contains("outside the allowable range"), e.Message);
            }
        }
        [Test]
        public void TestMethods()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.InsertChartRecord();
            //wb.dumpDrawingGroupRecords(true);
            //wb.dumpDrawingGroupRecords(false);
        }
        [Test]
        public void WriteProtection()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            Assert.IsFalse(wb.IsWriteProtected);

            wb.WriteProtectWorkbook("mypassword", "myuser");
            Assert.IsTrue(wb.IsWriteProtected);

            wb.UnwriteProtectWorkbook();
            Assert.IsFalse(wb.IsWriteProtected);
        }
        [Test]
        public void Bug50298()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("50298.xls");

            assertSheetOrder(wb, "Invoice", "Invoice1", "Digest", "Deferred", "Received");

            ISheet sheet = wb.CloneSheet(0);

            assertSheetOrder(wb, "Invoice", "Invoice1", "Digest", "Deferred", "Received", "Invoice (2)");

            wb.SetSheetName(wb.GetSheetIndex(sheet), "copy");

            assertSheetOrder(wb, "Invoice", "Invoice1", "Digest", "Deferred", "Received", "copy");

            wb.SetSheetOrder("copy", 0);

            assertSheetOrder(wb, "copy", "Invoice", "Invoice1", "Digest", "Deferred", "Received");

            wb.RemoveSheetAt(0);

            assertSheetOrder(wb, "Invoice", "Invoice1", "Digest", "Deferred", "Received");


            // check that the overall workbook serializes with its correct size
            int expected = wb.Workbook.Size;
            int written = wb.Workbook.Serialize(0, new byte[expected * 2]);

            Assert.AreEqual(expected, written, "Did not have the expected size when writing the workbook: written: " + written + ", but expected: " + expected);

            HSSFWorkbook read = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            assertSheetOrder(read, "Invoice", "Invoice1", "Digest", "Deferred", "Received");

            read.Close();
            wb.Close();
        }
        [Test]
        public void Bug50298a()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("50298.xls");

            assertSheetOrder(wb, "Invoice", "Invoice1", "Digest", "Deferred", "Received");

            ISheet sheet = wb.CloneSheet(0);

            assertSheetOrder(wb, "Invoice", "Invoice1", "Digest", "Deferred", "Received", "Invoice (2)");

            wb.SetSheetName(wb.GetSheetIndex(sheet), "copy");

            assertSheetOrder(wb, "Invoice", "Invoice1", "Digest", "Deferred", "Received", "copy");

            wb.SetSheetOrder("copy", 0);

            assertSheetOrder(wb, "copy", "Invoice", "Invoice1", "Digest", "Deferred", "Received");

            wb.RemoveSheetAt(0);

            assertSheetOrder(wb, "Invoice", "Invoice1", "Digest", "Deferred", "Received");

            wb.RemoveSheetAt(1);

            assertSheetOrder(wb, "Invoice", "Digest", "Deferred", "Received");

            wb.SetSheetOrder("Digest", 3);

            assertSheetOrder(wb, "Invoice", "Deferred", "Received", "Digest");

            // check that the overall workbook serializes with its correct size
            int expected = wb.Workbook.Size;
            int written = wb.Workbook.Serialize(0, new byte[expected * 2]);

            Assert.AreEqual(expected, written, "Did not have the expected size when writing the workbook: written: " + written + ", but expected: " + expected);

            HSSFWorkbook read = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            assertSheetOrder(wb, "Invoice", "Deferred", "Received", "Digest");
            read.Close();
            wb.Close();
        }

        [Test]
        public void Bug54500()
        {
            String nameName = "AName";
            String sheetName = "ASheet";
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("54500.xls");

            assertSheetOrder(wb, "Sheet1", "Sheet2", "Sheet3");

            wb.CreateSheet(sheetName);

            assertSheetOrder(wb, "Sheet1", "Sheet2", "Sheet3", "ASheet");

            IName n = wb.CreateName();
            n.NameName = (/*setter*/nameName);
            n.SheetIndex = (/*setter*/3);
            n.RefersToFormula = (/*setter*/sheetName + "!A1");

            assertSheetOrder(wb, "Sheet1", "Sheet2", "Sheet3", "ASheet");
            HSSFName name = wb.GetName(nameName) as HSSFName;
            Assert.IsNotNull(name);
            Assert.AreEqual("ASheet!A1", name.RefersToFormula);

            MemoryStream stream = new MemoryStream();
            wb.Write(stream, false);

            assertSheetOrder(wb, "Sheet1", "Sheet2", "Sheet3", "ASheet");
            Assert.AreEqual("ASheet!A1", name.RefersToFormula);

            wb.RemoveSheetAt(1);

            assertSheetOrder(wb, "Sheet1", "Sheet3", "ASheet");
            Assert.AreEqual("ASheet!A1", name.RefersToFormula);

            MemoryStream stream2 = new MemoryStream();
            wb.Write(stream2, false);

            assertSheetOrder(wb, "Sheet1", "Sheet3", "ASheet");
            Assert.AreEqual("ASheet!A1", name.RefersToFormula);

            HSSFWorkbook wb2 = new HSSFWorkbook(new ByteArrayInputStream(stream.ToArray()));
            ExpectName(wb2, nameName, "ASheet!A1");
            HSSFWorkbook wb3 = new HSSFWorkbook(new ByteArrayInputStream(stream2.ToArray()));
            ExpectName(wb3, nameName, "ASheet!A1");
            wb3.Close();
            wb2.Close();
            wb.Close();
        }

        private void ExpectName(HSSFWorkbook wb, String name, String expect)
        {
            HSSFName hssfName = wb.GetName(name) as HSSFName;
            Assert.IsNotNull(hssfName);
            Assert.AreEqual(expect, hssfName.RefersToFormula);
        }

        [Test]
        public void Best49423()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("49423.xls");

            bool found = false;
            int numSheets = workbook.NumberOfSheets;
            for (int i = 0; i < numSheets; i++)
            {
                HSSFSheet sheet = workbook.GetSheetAt(i) as HSSFSheet;
                IList<HSSFShape> shapes = (sheet.DrawingPatriarch as HSSFPatriarch).Children;
                foreach (HSSFShape shape in shapes)
                {
                    HSSFAnchor anchor = shape.Anchor;

                    if (anchor is HSSFClientAnchor)
                    {
                        // absolute coordinates
                        HSSFClientAnchor clientAnchor = (HSSFClientAnchor)anchor;
                        Assert.IsNotNull(clientAnchor);
                        //System.out.Println(clientAnchor.Row1 + "," + clientAnchor.Row2);
                        found = true;
                    }
                    else if (anchor is HSSFChildAnchor)
                    {
                        // shape is grouped and the anchor is expressed in the coordinate system of the group 
                        HSSFChildAnchor childAnchor = (HSSFChildAnchor)anchor;
                        Assert.IsNotNull(childAnchor);
                        //System.out.Println(childAnchor.Dy1 + "," + childAnchor.Dy2);
                        found = true;
                    }
                }
            }

            Assert.IsTrue(found, "Should find some images via Client or Child anchors, but did not find any at all");
            workbook.Close();
        }

        [Test]
        [Ignore("not found in poi 3.14")]
        public void Bug47245()
        {
            Assert.DoesNotThrow(() => HSSFTestDataSamples.OpenSampleWorkbook("47245_test.xls"));
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestRewriteFileBug58480()
        {
            FileInfo file = TempFile.CreateTempFile("TestHSSFWorkbook", ".xls");
            try
            {
                // create new workbook
                {
                    IWorkbook workbook = new HSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("foo");
                    IRow row = sheet.CreateRow(1);
                    row.CreateCell(1).SetCellValue("bar");

                    WriteAndCloseWorkbook(workbook, file);
                }

                // edit the workbook
                {
                    NPOIFSFileSystem fs = new NPOIFSFileSystem(file, false);
                    try
                    {
                        DirectoryNode root = fs.Root;
                        IWorkbook workbook = new HSSFWorkbook(root, true);
                        ISheet sheet = workbook.GetSheet("foo");
                        sheet.GetRow(1).CreateCell(2).SetCellValue("baz");

                        WriteAndCloseWorkbook(workbook, file);
                    }
                    finally
                    {
                        fs.Close();
                    }
                }
            }
            finally
            {
                Assert.IsTrue(file.Exists);
                file.Delete();
                Assert.IsTrue(!File.Exists(file.FullName));
            }
        }

        private void WriteAndCloseWorkbook(IWorkbook workbook, FileInfo file)
        {
            ByteArrayOutputStream bytesOut = new ByteArrayOutputStream();
            workbook.Write(bytesOut, false);
            workbook.Close();

            byte[] byteArray = bytesOut.ToByteArray();
            bytesOut.Close();

            FileStream fileOut = new FileStream(file.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fileOut.Write(byteArray, 0, byteArray.Length);
            fileOut.Close();

        }


        [Test]
        public void CloseDoesNotModifyWorkbook()
        {
            String filename = "SampleSS.xls";
            FileInfo file = POIDataSamples.GetSpreadSheetInstance().GetFileInfo(filename);
            IWorkbook wb;

            // File via POIFileStream (java.io)
            wb = new HSSFWorkbook(new POIFSFileSystem(file));
            assertCloseDoesNotModifyFile(filename, wb);

            // File via NPOIFileStream (java.nio)
            wb = new HSSFWorkbook(new NPOIFSFileSystem(file));
            assertCloseDoesNotModifyFile(filename, wb);

            // InputStream
            wb = new HSSFWorkbook(file.OpenRead());
            assertCloseDoesNotModifyFile(filename, wb);
        }

        [Test]
        public void SetSheetOrderToEnd()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            workbook.CreateSheet("A");
            try
            {
                for (int i = 0; i < 2 * workbook.InternalWorkbook.Records.Count; i++)
                {
                    workbook.SetSheetOrder("A", 0);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Moving a sheet to the end should not throw an exception, but threw ", e);
            }
        }


        [Test]
        public void InvalidInPlaceWrite()
        {
            HSSFWorkbook wb;

            // Can't work for new files
            wb = new HSSFWorkbook();
            try
            {
                wb.Write();
                Assert.Fail("Shouldn't work for new files");
            }
            catch (InvalidOperationException)
            {
                // expected here
            }
            wb.Close();

            // Can't work for InputStream opened files
            wb = new HSSFWorkbook(
                POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("SampleSS.xls"));
            try
            {
                wb.Write();
                Assert.Fail("Shouldn't work for InputStream");
            }
            catch (InvalidOperationException)
            {
                // expected here
            }
            wb.Close();

            // Can't work for OPOIFS
            OPOIFSFileSystem ofs = new OPOIFSFileSystem(
                    POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("SampleSS.xls"));
            wb = new HSSFWorkbook(ofs.Root, true);
            try
            {
                wb.Write();
                Assert.Fail("Shouldn't work for OPOIFSFileSystem");
            }
            catch (InvalidOperationException)
            {
                // expected here
            }
            wb.Close();

            // Can't work for Read-Only files
            NPOIFSFileSystem fs = new NPOIFSFileSystem(
                    POIDataSamples.GetSpreadSheetInstance().GetFile("SampleSS.xls"), true);
            wb = new HSSFWorkbook(fs);
            try
            {
                wb.Write();
                Assert.Fail("Shouldn't work for Read Only");
            }
            catch (InvalidOperationException)
            {
                // expected here
            }
            wb.Close();
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void InPlaceWrite()
        {
            // Setup as a copy of a known-good file
            FileInfo file = TempFile.CreateTempFile("TestHSSFWorkbook", ".xls");
            Stream inputStream = POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("SampleSS.xls");
            try
            {
                Stream outputStream = file.Open(FileMode.Open, FileAccess.ReadWrite);
                try
                {
                    IOUtils.Copy(inputStream, outputStream);
                }
                finally
                {
                    outputStream.Close();
                }
            }
            finally
            {
                inputStream.Close();
            }

            // Open from the temp file in read-write mode
            HSSFWorkbook wb = new HSSFWorkbook(new NPOIFSFileSystem(file, false));
            Assert.AreEqual(3, wb.NumberOfSheets);

            // Change
            wb.RemoveSheetAt(2);
            wb.RemoveSheetAt(1);
            wb.GetSheetAt(0).GetRow(0).GetCell(0).SetCellValue("Changed!");

            // Save in-place, close, re-open and check
            wb.Write();
            wb.Close();

            wb = new HSSFWorkbook(new NPOIFSFileSystem(file));
            Assert.AreEqual(1, wb.NumberOfSheets);
            Assert.AreEqual("Changed!", wb.GetSheetAt(0).GetRow(0).GetCell(0).ToString());

            wb.Close();
        }

        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestWriteToNewFile()
        {
            // Open from a Stream
            HSSFWorkbook wb = new HSSFWorkbook(
                    POIDataSamples.GetSpreadSheetInstance().OpenResourceAsStream("SampleSS.xls"));
            // Save to a new temp file
            FileInfo file = TempFile.CreateTempFile("TestHSSFWorkbook", ".xls");
            wb.Write(file);
            wb.Close();

            // Read and check
            wb = new HSSFWorkbook(new NPOIFSFileSystem(file));
            Assert.AreEqual(3, wb.NumberOfSheets);
            wb.Close();
        }


        [Test]
        public void TestBug854()
        {
            Assert.DoesNotThrow(() => HSSFTestDataSamples.OpenSampleWorkbook("ATM.xls"));
        }
    }
}
