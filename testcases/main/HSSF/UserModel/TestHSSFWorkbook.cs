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
    /**
     *
     */
    [TestFixture]
    public class TestHSSFWorkbook : BaseTestWorkbook
    {
        public TestHSSFWorkbook()
            : base(HSSFITestDataProvider.Instance)
        { }
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

        [Test]
        public void CaseInsensitiveNames()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet originalSheet = b.CreateSheet("Sheet1");
            NPOI.SS.UserModel.ISheet fetchedSheet = b.GetSheet("sheet1");
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
        public void WindowOneDefaults()
        {
            HSSFWorkbook b = new HSSFWorkbook();
            try
            {
                Assert.AreEqual(b.ActiveSheetIndex, 0);
                Assert.AreEqual(b.FirstVisibleTab, 0);
            }
            catch (NullReferenceException)
            {
                Assert.Fail("WindowOneRecord in Workbook is probably not initialized");
            }
        }
        [Test]
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
        public void SheetClone()
        {
            // First up, try a simple file
            HSSFWorkbook b = new HSSFWorkbook();
            Assert.AreEqual(0, b.NumberOfSheets);
            b.CreateSheet("Sheet One");
            b.CreateSheet("Sheet Two");

            Assert.AreEqual(2, b.NumberOfSheets);
            b.CloneSheet(0);
            Assert.AreEqual(3, b.NumberOfSheets);

            // Now try a problem one with drawing records in it
            b = OpenSample("SheetWithDrawing.xls");
            Assert.AreEqual(1, b.NumberOfSheets);
            b.CloneSheet(0);
            Assert.AreEqual(2, b.NumberOfSheets);
        }
        [Test]
        public void ReadWriteWithCharts()
        {
            HSSFWorkbook b;
            NPOI.SS.UserModel.ISheet s;

            // Single chart, two sheets
            b = OpenSample("44010-SingleChart.xls");
            Assert.AreEqual(2, b.NumberOfSheets);
            Assert.AreEqual("Graph2", b.GetSheetName(1));
            s = b.GetSheetAt(1);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);

            // Has chart on 1st sheet??
            // FIXME
            Assert.IsNotNull(b.GetSheetAt(0).DrawingPatriarch);
            Assert.IsNull(b.GetSheetAt(1).DrawingPatriarch);
            Assert.IsFalse(((HSSFPatriarch)b.GetSheetAt(0).DrawingPatriarch).ContainsChart());

            // We've now called DrawingPatriarch so
            //  everything will be all screwy
            // So, start again
            b = OpenSample("44010-SingleChart.xls");

            b = WriteRead(b);
            Assert.AreEqual(2, b.NumberOfSheets);
            s = b.GetSheetAt(1);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);


            // Two charts, three sheets
            b = OpenSample("44010-TwoCharts.xls");
            Assert.AreEqual(3, b.NumberOfSheets);

            s = b.GetSheetAt(1);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
            s = b.GetSheetAt(2);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);

            // Has chart on 1st sheet??
            // FIXME
            Assert.IsNotNull(b.GetSheetAt(0).DrawingPatriarch);
            Assert.IsNull(b.GetSheetAt(1).DrawingPatriarch);
            Assert.IsNull(b.GetSheetAt(2).DrawingPatriarch);
            Assert.IsFalse(((HSSFPatriarch)b.GetSheetAt(0).DrawingPatriarch).ContainsChart());

            // We've now called DrawingPatriarch so
            //  everything will be all screwy
            // So, start again
            b = OpenSample("44010-TwoCharts.xls");

            b = WriteRead(b);
            Assert.AreEqual(3, b.NumberOfSheets);

            s = b.GetSheetAt(1);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
            s = b.GetSheetAt(2);
            Assert.AreEqual(0, s.FirstRowNum);
            Assert.AreEqual(8, s.LastRowNum);
        }

        private static HSSFWorkbook WriteRead(HSSFWorkbook b)
        {
            return HSSFTestDataSamples.WriteOutAndReadBack(b);
        }

        [Test]
        public void SelectedSheet_bug44523()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet1 = wb.CreateSheet("Sheet1");
            NPOI.SS.UserModel.ISheet sheet2 = wb.CreateSheet("Sheet2");
            NPOI.SS.UserModel.ISheet sheet3 = wb.CreateSheet("Sheet3");
            NPOI.SS.UserModel.ISheet sheet4 = wb.CreateSheet("Sheet4");

            ConfirmActiveSelected(sheet1, true);
            ConfirmActiveSelected(sheet2, false);
            ConfirmActiveSelected(sheet3, false);
            ConfirmActiveSelected(sheet4, false);

            wb.SetSelectedTab(1);
            // see Javadoc, in this case selected means "active"
            Assert.AreEqual(wb.ActiveSheetIndex, (short)wb.ActiveSheetIndex);


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

        //public void SelectMultiple()
        //{
        //    HSSFWorkbook wb = new HSSFWorkbook();
        //    NPOI.SS.UserModel.Sheet sheet1 = wb.CreateSheet("Sheet1");
        //    NPOI.SS.UserModel.Sheet sheet2 = wb.CreateSheet("Sheet2");
        //    NPOI.SS.UserModel.Sheet sheet3 = wb.CreateSheet("Sheet3");
        //    NPOI.SS.UserModel.Sheet sheet4 = wb.CreateSheet("Sheet4");
        //    NPOI.SS.UserModel.Sheet sheet5 = wb.CreateSheet("Sheet5");
        //    NPOI.SS.UserModel.Sheet sheet6 = wb.CreateSheet("Sheet6");

        //    wb.SetSelectedTabs(new int[] { 0, 2, 3 });

        //    Assert.AreEqual(true, sheet1.IsSelected);
        //    Assert.AreEqual(false, sheet2.IsSelected);
        //    Assert.AreEqual(true, sheet3.IsSelected);
        //    Assert.AreEqual(true, sheet4.IsSelected);
        //    Assert.AreEqual(false, sheet5.IsSelected);
        //    Assert.AreEqual(false, sheet6.IsSelected);

        //    wb.SetSelectedTabs(new int[] { 1, 3, 5 });

        //    Assert.AreEqual(false, sheet1.IsSelected);
        //    Assert.AreEqual(true, sheet2.IsSelected);
        //    Assert.AreEqual(false, sheet3.IsSelected);
        //    Assert.AreEqual(true, sheet4.IsSelected);
        //    Assert.AreEqual(false, sheet5.IsSelected);
        //    Assert.AreEqual(true, sheet6.IsSelected);

        //    Assert.AreEqual(true, sheet1.IsActive);
        //    Assert.AreEqual(false, sheet2.IsActive);


        //    Assert.AreEqual(true, sheet1.IsActive);
        //    Assert.AreEqual(false, sheet3.IsActive);
        //    wb.SetActiveSheet(2);
        //    Assert.AreEqual(false, sheet1.IsActive);
        //    Assert.AreEqual(true, sheet3.IsActive);

        //    if (false)
        //    { // helpful if viewing this workbook in excel:
        //        sheet1.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Sheet1"));
        //        sheet2.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Sheet2"));
        //        sheet3.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Sheet3"));
        //        sheet4.CreateRow(0).CreateCell(0).SetCellValue(new HSSFRichTextString("Sheet4"));

        //        try
        //        {
        //            File fOut = TempFile.CreateTempFile("sheetMultiSelect", ".xls");
        //            FileOutputStream os = new FileOutputStream(fOut);
        //            wb.Write(os);
        //            os.Close();
        //        }
        //        catch (IOException e)
        //        {
        //            throw new RuntimeException(e);
        //        }
        //    }
        //}

        [Test]
        public void ActiveSheetAfterDelete_bug40414()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet0 = wb.CreateSheet("Sheet0");
            NPOI.SS.UserModel.ISheet sheet1 = wb.CreateSheet("Sheet1");
            NPOI.SS.UserModel.ISheet sheet2 = wb.CreateSheet("Sheet2");
            NPOI.SS.UserModel.ISheet sheet3 = wb.CreateSheet("Sheet3");
            NPOI.SS.UserModel.ISheet sheet4 = wb.CreateSheet("Sheet4");

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

        private static void ConfirmActiveSelected(NPOI.SS.UserModel.ISheet sheet, bool expected)
        {
            ConfirmActiveSelected(sheet, expected, expected);
        }


        private static void ConfirmActiveSelected(NPOI.SS.UserModel.ISheet sheet,
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
         * Checks that us and NPOI.SS.UserModel.Name play nicely with named ranges
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
            NPOI.SS.UserModel.IName n;

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
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("TestRRaC.xls");
            NameRecord nr;
            Assert.AreEqual(3, wb.Workbook.NumNames);
            nr = wb.Workbook.GetNameRecord(2);
            // TODO - render full row and full column refs properly
            Assert.AreEqual("Sheet2!$A$1:$IV$1", HSSFFormulaParser.ToFormulaString(wb, nr.NameDefinition)); // 1:1

            try
            {
                wb.GetSheetAt(3).RepeatingRows = (CellRangeAddress.ValueOf("9:12"));
                wb.GetSheetAt(3).RepeatingColumns = (CellRangeAddress.ValueOf("E:F"));
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Builtin (7) already exists for sheet (4)"))
                {
                    // there was a problem in the code which locates the existing print titles name record 
                    throw new Exception("Identified bug 45720b");
                }
                throw e;
            }
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.AreEqual(3, wb.Workbook.NumNames);
            nr = wb.Workbook.GetNameRecord(2);
            Assert.AreEqual("Sheet2!E:F,Sheet2!$A$9:$IV$12", HSSFFormulaParser.ToFormulaString(wb, nr.NameDefinition)); // E:F,9:12
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
            wb.Write(out1);
            byte[] bytes = out1.ToArray();
            POIFSFileSystem fs2 = new POIFSFileSystem(new MemoryStream(bytes));
            ClassID clsid2 = fs2.Root.StorageClsid;

            Assert.IsTrue(clsid1.Equals(clsid2));
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
                for (int i = 0; i < objects.Count; i++)
                {
                    HSSFObjectData embeddedObject = objects[i];
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
        }

        [Test]
        public void ChangeSheetNameWithSharedFormulas()
        {
            ChangeSheetNameWithSharedFormulas("shared_formulas.xls");
        }
        [Test]
        public void EmptyDirectoryNode()
        {
            try
            {
                Assert.IsNotNull(new HSSFWorkbook(new POIFSFileSystem()));
                Assert.Fail("Should catch exception about invalid POIFSFileSystem");
            }
            catch (ArgumentException e)
            {
                Assert.IsTrue(e.Message.Contains("does not contain a BIFF8"), e.Message);
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

            // see Javadoc, in this case selected means "active"
            Assert.AreEqual(wb.ActiveSheetIndex, (short)wb.ActiveSheetIndex);

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
            //wb.DisplayedTab=((short)2);
            //Assert.AreEqual(2, wb.FirstVisibleTab);
            //Assert.AreEqual(2, wb.DisplayedTab);
        }
        [Test]
        public void AddSheetTwice()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet1 = (HSSFSheet)wb.CreateSheet("Sheet1");
            Assert.IsNotNull(sheet1);
            try
            {
                wb.CreateSheet("Sheet1");
                Assert.Fail("Should fail if we add the same sheet twice");
            }
            catch (ArgumentException)
            {
                //Assert.IsTrue(e.Message.Contains("already Contains a sheet of this name"), e.Message);
            }
        }
        [Test]
        public void GetSheetIndex()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet1 = (HSSFSheet)wb.CreateSheet("Sheet1");
            HSSFSheet sheet2 = (HSSFSheet)wb.CreateSheet("Sheet2");
            HSSFSheet sheet3 = (HSSFSheet)wb.CreateSheet("Sheet3");
            HSSFSheet sheet4 = (HSSFSheet)wb.CreateSheet("Sheet4");

            Assert.AreEqual(0, wb.GetSheetIndex(sheet1));
            Assert.AreEqual(1, wb.GetSheetIndex(sheet2));
            Assert.AreEqual(2, wb.GetSheetIndex(sheet3));
            Assert.AreEqual(3, wb.GetSheetIndex(sheet4));

            // remove sheets
            wb.RemoveSheetAt(0);
            wb.RemoveSheetAt(2);

            // ensure that sheets are Moved up and Removed sheets are not found any more
            Assert.AreEqual(-1, wb.GetSheetIndex(sheet1));
            Assert.AreEqual(0, wb.GetSheetIndex(sheet2));
            Assert.AreEqual(1, wb.GetSheetIndex(sheet3));
            Assert.AreEqual(-1, wb.GetSheetIndex(sheet4));
        }
        [Test]
        public void ExternSheetIndex()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            wb.CreateSheet("Sheet1");
            wb.CreateSheet("Sheet2");

            Assert.AreEqual(0, wb.GetExternalSheetIndex(0));
            Assert.AreEqual(1, wb.GetExternalSheetIndex(1));

            //The following methods are obsoleted

            //Assert.AreEqual("Sheet1", wb.FindSheetNameFromExternSheet(0));
            //Assert.AreEqual("Sheet2", wb.FindSheetNameFromExternSheet(1));
            ////Assert.AreEqual(null, wb.FindSheetNameFromExternSheet(2));

            //Assert.AreEqual(0, wb.GetSheetIndexFromExternSheetIndex(0));
            //Assert.AreEqual(1, wb.GetSheetIndexFromExternSheetIndex(1));
            //Assert.AreEqual(-1, wb.GetSheetIndexFromExternSheetIndex(2));
            //Assert.AreEqual(-1, wb.GetSheetIndexFromExternSheetIndex(100));
        }
        [Test]
        public void SSTString()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            int sst = wb.AddSSTString("somestring");
            Assert.AreEqual("somestring", wb.GetSSTString(sst));
            ////Assert.IsNull(wb.GetSSTString(sst+1));
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
            Assert.AreEqual("ASheet!A1", wb.GetName(nameName).RefersToFormula);

            MemoryStream stream = new MemoryStream();
            wb.Write(stream);

            assertSheetOrder(wb, "Sheet1", "Sheet2", "Sheet3", "ASheet");
            Assert.AreEqual("ASheet!A1", wb.GetName(nameName).RefersToFormula);

            wb.RemoveSheetAt(1);

            assertSheetOrder(wb, "Sheet1", "Sheet3", "ASheet");
            Assert.AreEqual("ASheet!A1", wb.GetName(nameName).RefersToFormula);

            MemoryStream stream2 = new MemoryStream();
            wb.Write(stream2);

            assertSheetOrder(wb, "Sheet1", "Sheet3", "ASheet");
            Assert.AreEqual("ASheet!A1", wb.GetName(nameName).RefersToFormula);

            expectName(
                    new HSSFWorkbook(new MemoryStream(stream.ToArray())),
                    nameName, "ASheet!A1");
            expectName(
                    new HSSFWorkbook(
                            new MemoryStream(stream2.ToArray())),
                    nameName, "ASheet!A1");
        }

        private void expectName(HSSFWorkbook wb, String name, String expect)
        {
            Assert.AreEqual(expect, wb.GetName(name).RefersToFormula);
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
        }

    }
}
