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
    /**
     *
     */
    [TestFixture]
    public class TestHSSFWorkbook:BaseTestWorkbook
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
        public void TestCaseInsensitiveNames()
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
        public void TestDuplicateNames()
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
        public void TestWindowOneDefaults()
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
        public void TestSheetClone()
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
        public void TestReadWriteWithCharts()
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
        public void TestSelectedSheet_bug44523()
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

        //public void TestSelectMultiple()
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
        public void TestActiveSheetAfterDelete_bug40414()
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
        public void TestSheetSerializeSizeMisMatch_bug45066()
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
        public void TestNamesToDeleteSheets()
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
        public void TestExtraDataAfterEOFRecord()
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
        public void TestFindBuiltInNameRecord()
        {
            // TestRRaC has multiple (3) built-in name records
            // The second print titles name record has SheetNumber==4
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("TestRRaC.xls");
            NameRecord nr;
            Assert.AreEqual(3, wb.Workbook.NumNames);
            nr = wb.Workbook.GetNameRecord(2);
            // TODO - render full row and full column refs properly
            Assert.AreEqual("Sheet2!$A$1:$IV$1", HSSFFormulaParser.ToFormulaString(wb,nr.NameDefinition)); // 1:1

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
            Assert.AreEqual("Sheet2!E:F,Sheet2!$A$9:$IV$12", HSSFFormulaParser.ToFormulaString(wb,nr.NameDefinition)); // E:F,9:12
        }
        /**
     * Test that the storage clsid property is preserved
     */
        [Test]
        public void Test47920()
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
         * Tests that we can work with both {@link POIFSFileSystem}
         *  and {@link NPOIFSFileSystem}
         */
        [Test]
        public void TestDifferentPOIFS()
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
        public void TestWordDocEmbeddedInXls()
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
        public void TestWriteWorkbookFromNPOIFS()
        {
            //throw new NotImplementedException("class NPOIFSFileSystem is not implemented");
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("WithEmbeddedObjects.xls");
            NPOIFSFileSystem fs = new NPOIFSFileSystem(is1);

            // Start as NPOIFS
            HSSFWorkbook wb = new HSSFWorkbook(fs.Root, true);
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.AreEqual("Root xls", wb.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);

            // Will switch to POIFS
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.AreEqual(3, wb.NumberOfSheets);
            Assert.AreEqual("Root xls", wb.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
        }

        [Test]
        public void TestCellStylesLimit()
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
        public void TestSetSheetOrderHSSF()
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
                    ComparisonOperator.BETWEEN, "'first sheet'!D1", "'other sheet'!D1");

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
        public void TestClonePictures()
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
        public void TestChangeSheetNameWithSharedFormulas()
        {
            ChangeSheetNameWithSharedFormulas("shared_formulas.xls");
        }
    }
}