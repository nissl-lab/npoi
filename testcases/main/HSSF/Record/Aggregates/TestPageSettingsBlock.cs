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

using TestCases.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using NUnit.Framework;
using NPOI.Util;
using NPOI.HSSF.Model;
using TestCases.HSSF.UserModel;
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.Aggregates;
namespace TestCases.HSSF.Record.Aggregates
{

    /**
     * Tess for {@link PageSettingsBlock}
     *
     * @author Dmitriy Kumshayev
     */
    [TestFixture]
    public class TestPageSettingsBlock
    {
        [Test]
        public void TestPrintSetup_bug46548()
        {

            // PageSettingBlock in this file Contains PLS (sid=x004D) record
            // followed by ContinueRecord (sid=x003C)
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex46548-23133.xls");
            ISheet sheet = wb.GetSheetAt(0);
            IPrintSetup ps = sheet.PrintSetup;

            try
            {
                int copies = ps.Copies;
            }
            catch (NullReferenceException)
            {
                Assert.Fail("Identified bug 46548: PageSettingBlock missing PrintSetupRecord record");
            }
        }

        /**
         * Bug 46840 occurred because POI failed to recognise HEADERFOOTER as part of the
         * {@link PageSettingsBlock}.
         */
        [Test]
        public void TestHeaderFooter_bug46840()
        {

            int rowIx = 5;
            int colIx = 6;
            NumberRecord nr = new NumberRecord();
            nr.Row = (rowIx);
            nr.Column = ((short)colIx);
            nr.Value = (3.0);

            NPOI.HSSF.Record.Record[] recs = {
                BOFRecord.CreateSheetBOF(),
                new HeaderRecord("&LSales Figures"),
                new FooterRecord("&LJanuary"),
                new HeaderFooterRecord(HexRead.ReadFromString("9C 08 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 C4 60 00 00 00 00 00 00 00 00")),
                new DimensionsRecord(),
                new WindowTwoRecord(),
                new UserSViewBegin(HexRead.ReadFromString("ED 77 3B 86 BC 3F 37 4C A9 58 60 23 43 68 54 4B 01 00 00 00 64 00 00 00 40 00 00 00 02 00 00 00 3D 80 04 00 00 00 00 00 00 00 0C 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 F0 3F FF FF 01 00")),
                new HeaderRecord("&LSales Figures"),
                new FooterRecord("&LJanuary"),
                new HeaderFooterRecord(HexRead.ReadFromString("9C 08 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 C4 60 00 00 00 00 00 00 00 00")),
                new UserSViewEnd(HexRead.ReadFromString("01, 00")),

                EOFRecord.instance,
        };
            RecordStream rs = new RecordStream(Arrays.AsList(recs), 0);
            InternalSheet sheet;
            try
            {
                sheet = InternalSheet.CreateSheet(rs);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("two Page Settings Blocks found in the same sheet"))
                {
                    throw new AssertionException("Identified bug 46480");
                }
                throw;
            }

            TestCases.HSSF.UserModel.RecordInspector.RecordCollector rv = new TestCases.HSSF.UserModel.RecordInspector.RecordCollector();
            sheet.VisitContainedRecords(rv, rowIx);
            NPOI.HSSF.Record.Record[] outRecs = rv.Records;
            Assert.AreEqual(13, outRecs.Length);
        }

        /**
         * Bug 46953 occurred because POI didn't handle late PSB records properly.
         */
        [Test]
        public void TestLateHeaderFooter_bug46953()
        {

            int rowIx = 5;
            int colIx = 6;
            NumberRecord nr = new NumberRecord();
            nr.Row = (rowIx);
            nr.Column = ((short)colIx);
            nr.Value = (3.0);

            NPOI.HSSF.Record.Record[] recs = {
                BOFRecord.CreateSheetBOF(),
                new HeaderRecord("&LSales Figures"),
                new FooterRecord("&LJanuary"),
                new DimensionsRecord(),
                new WindowTwoRecord(),
                new HeaderFooterRecord(HexRead.ReadFromString("9C 08 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 C4 60 00 00 00 00 00 00 00 00")),
                EOFRecord.instance,
        };
            RecordStream rs = new RecordStream(Arrays.AsList(recs), 0);
            InternalSheet sheet = InternalSheet.CreateSheet(rs);

            RecordInspector.RecordCollector rv = new RecordInspector.RecordCollector();
            sheet.VisitContainedRecords(rv, 0);
            NPOI.HSSF.Record.Record[] outRecs = rv.Records;
            if (outRecs[4] == EOFRecord.instance)
            {
                throw new AssertionException("Identified bug 46953 - EOF incorrectly Appended to PSB");
            }
            Assert.AreEqual(recs.Length + 1, outRecs.Length); // +1 for index record

            Assert.AreEqual(typeof(BOFRecord), outRecs[0].GetType());
            Assert.AreEqual(typeof(IndexRecord), outRecs[1].GetType());
            Assert.AreEqual(typeof(HeaderRecord), outRecs[2].GetType());
            Assert.AreEqual(typeof(FooterRecord), outRecs[3].GetType());
            Assert.AreEqual(typeof(HeaderFooterRecord), outRecs[4].GetType());
            Assert.AreEqual(typeof(DimensionsRecord), outRecs[5].GetType());
            Assert.AreEqual(typeof(WindowTwoRecord), outRecs[6].GetType());
            Assert.AreEqual(typeof(EOFRecord), outRecs[7].GetType());
        }
        /**
         * Bug 47199 was due to the margin records being located well after the Initial PSB records.
         * The example file supplied (attachment 23710) had three non-PSB record types
         * between the PRINTSETUP record and first MARGIN record:
         * <ul>
         * <li>PRINTSETUP(0x00A1)</li>
         * <li>DEFAULTCOLWIDTH(0x0055)</li>
         * <li>COLINFO(0x007D)</li>
         * <li>DIMENSIONS(0x0200)</li>
         * <li>BottomMargin(0x0029)</li>
         * </ul>
         */
        [Test]
        public void TestLateMargins_bug47199()
        {

            NPOI.HSSF.Record.Record[] recs = {
                BOFRecord.CreateSheetBOF(),
                new HeaderRecord("&LSales Figures"),
                new FooterRecord("&LJanuary"),
                new DimensionsRecord(),
                CreateBottomMargin(0.787F),
                new WindowTwoRecord(),
                EOFRecord.instance,
        };
            RecordStream rs = new RecordStream(Arrays.AsList(recs), 0);

            InternalSheet sheet;
            try
            {
                sheet = InternalSheet.CreateSheet(rs);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("two Page Settings Blocks found in the same sheet"))
                {
                    throw new AssertionException("Identified bug 47199a - failed to process late margings records");
                }
                throw;
            }

            TestCases.HSSF.UserModel.RecordInspector.RecordCollector rv = new TestCases.HSSF.UserModel.RecordInspector.RecordCollector();
            sheet.VisitContainedRecords(rv, 0);
            NPOI.HSSF.Record.Record[] outRecs = rv.Records;
            Assert.AreEqual(recs.Length + 1, outRecs.Length); // +1 for index record

            Assert.AreEqual(typeof(BOFRecord), outRecs[0].GetType());
            Assert.AreEqual(typeof(IndexRecord), outRecs[1].GetType());
            Assert.AreEqual(typeof(HeaderRecord), outRecs[2].GetType());
            Assert.AreEqual(typeof(FooterRecord), outRecs[3].GetType());
            Assert.AreEqual(typeof(DimensionsRecord), outRecs[5].GetType());
            Assert.AreEqual(typeof(WindowTwoRecord), outRecs[6].GetType());
            Assert.AreEqual(typeof(EOFRecord), outRecs[7].GetType());
        }

        private NPOI.HSSF.Record.Record CreateBottomMargin(float value)
        {
            BottomMarginRecord result = new BottomMarginRecord();
            result.Margin = (value);
            return result;
        }

        /**
         * The PageSettingsBlock should not allow multiple copies of the same record.  This extra assertion
         * was Added while fixing bug 47199.  All existing POI Test samples comply with this requirement.
         */
        [Test]
        public void TestDuplicatePSBRecord_bug47199()
        {

            // Hypothetical Setup of PSB records which should cause POI to crash
            NPOI.HSSF.Record.Record[] recs = {
                new HeaderRecord("&LSales Figures"),
                new HeaderRecord("&LInventory"),
        };
            RecordStream rs = new RecordStream(Arrays.AsList(recs), 0);

            try
            {
                new PageSettingsBlock(rs);
                throw new AssertionException("Identified bug 47199b - duplicate PSB records should not be allowed");
            }
            catch (RecordFormatException e)
            {
                if (e.Message.Equals("Duplicate PageSettingsBlock record (sid=0x14)"))
                {
                    // expected during successful Test
                }
                else
                {
                    throw new AssertionException("Expected RecordFormatException due to duplicate PSB record");
                }
            }
        }

        private static UnknownRecord ur(int sid, String hexData)
        {
            return new UnknownRecord(sid, HexRead.ReadFromString(hexData));
        }

        /**
         * Excel tolerates missing header / footer records, but Adds them (empty) in when re-saving.
         * This is not critical functionality but it has been decided to keep POI consistent with
         * Excel in this regard.
         */
        [Test]
        public void TestMissingHeaderFooter()
        {
            // Initialise PSB with some records, but not the header / footer
            NPOI.HSSF.Record.Record[] recs = {
                new HCenterRecord(),
                new VCenterRecord(),
        };
            RecordStream rs = new RecordStream(Arrays.AsList(recs), 0);
            PageSettingsBlock psb = new PageSettingsBlock(rs);

            // serialize the PSB to see what records come out
            RecordInspector.RecordCollector rc = new RecordInspector.RecordCollector();
            psb.VisitContainedRecords(rc);
            NPOI.HSSF.Record.Record[] outRecs = rc.Records;

            if (outRecs.Length == 2)
            {
                throw new AssertionException("PageSettingsBlock didn't add missing header/footer records");
            }
            Assert.AreEqual(4, outRecs.Length);
            Assert.AreEqual(typeof(HeaderRecord), outRecs[0].GetType());
            Assert.AreEqual(typeof(FooterRecord), outRecs[1].GetType());
            Assert.AreEqual(typeof(HCenterRecord), outRecs[2].GetType());
            Assert.AreEqual(typeof(VCenterRecord), outRecs[3].GetType());

            // make sure the Added header / footer records are empty
            HeaderRecord hr = (HeaderRecord)outRecs[0];
            Assert.AreEqual("", hr.Text);
            FooterRecord fr = (FooterRecord)outRecs[1];
            Assert.AreEqual("", fr.Text);
        }

        /**
         * Apparently it's OK to have more than one PLS record.
         * Attachment 23866 from bug 47415 had a PageSettingsBlock with two PLS records.  This file
         * seems to open OK in Excel(2007) but both PLS records are Removed (perhaps because the
         * specified printers were not available on the Testing machine).  Since the example file does
         * not upset Excel, POI will preserve multiple PLS records.</p>
         *
         * As of June 2009, PLS is still uninterpreted by POI
         */
        [Test]
        public void TestDuplicatePLS_bug47415()
        {
            NPOI.HSSF.Record.Record plsA = ur(UnknownRecord.PLS_004D, "BA AD F0 0D");
            NPOI.HSSF.Record.Record plsB = ur(UnknownRecord.PLS_004D, "DE AD BE EF");
            NPOI.HSSF.Record.Record contB1 = new ContinueRecord(HexRead.ReadFromString("FE ED"));
            NPOI.HSSF.Record.Record contB2 = new ContinueRecord(HexRead.ReadFromString("FA CE"));
            NPOI.HSSF.Record.Record[] recs = {
                new HeaderRecord("&LSales Figures"),
                new FooterRecord("&LInventory"),
                new HCenterRecord(),
                new VCenterRecord(),
                plsA,
                plsB, contB1, contB2, // make sure continuing PLS is still OK
		};
            RecordStream rs = new RecordStream(Arrays.AsList(recs), 0);
            PageSettingsBlock psb;
            try
            {
                psb = new PageSettingsBlock(rs);
            }
            catch (RecordFormatException e)
            {
                if ("Duplicate PageSettingsBlock record (sid=0x4d)".Equals(e.Message))
                {
                    throw new AssertionException("Identified bug 47415");
                }
                throw e;
            }

            // serialize the PSB to see what records come out
            RecordInspector.RecordCollector rc = new RecordInspector.RecordCollector();
            psb.VisitContainedRecords(rc);
            NPOI.HSSF.Record.Record[] outRecs = rc.Records;

            // records were assembled in standard order, so this simple check is OK
            Assert.IsTrue(Arrays.Equals(recs, outRecs));
        }
        [Test]
        public void TestDuplicateHeaderFooter_bug48026()
        {

            NPOI.HSSF.Record.Record[] recs = {
                BOFRecord.CreateSheetBOF(),
                new IndexRecord(),

                //PageSettingsBlock
                new HeaderRecord("&LDecember"),
                new FooterRecord("&LJanuary"),
                new DimensionsRecord(),

                new WindowTwoRecord(),

                //CustomViewSettingsRecordAggregate
                new UserSViewBegin(HexRead.ReadFromString("53 CE BD CC DE 38 44 45 97 C1 5C 89 F9 37 32 1B 01 00 00 00 64 00 00 00 40 00 00 00 03 00 00 00 7D 00 00 20 00 00 34 00 00 00 18 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF FF FF FF")),
                new SelectionRecord(0, 0),
                new UserSViewEnd(HexRead.ReadFromString("01 00")),

                // two HeaderFooterRecord records, the first one has zero GUID (16 bytes at offset 12) and belongs to the PSB,
                // the other is matched with a CustomViewSettingsRecordAggregate having UserSViewBegin with the same GUID
                new HeaderFooterRecord(HexRead.ReadFromString("9C 08 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 34 33 00 00 00 00 00 00 00 00")),
                new HeaderFooterRecord(HexRead.ReadFromString("9C 08 00 00 00 00 00 00 00 00 00 00 53 CE BD CC DE 38 44 45 97 C1 5C 89 F9 37 32 1B 34 33 00 00 00 00 00 00 00 00")),

                EOFRecord.instance,
            };
            RecordStream rs = new RecordStream(Arrays.AsList(recs), 0);
            InternalSheet sheet;
            try
            {
                sheet = InternalSheet.CreateSheet(rs);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Duplicate PageSettingsBlock record (sid=0x89c)"))
                {
                    throw new AssertionException("Identified bug 48026");
                }
                throw;
            }

            RecordInspector.RecordCollector rv = new RecordInspector.RecordCollector();
            sheet.VisitContainedRecords(rv, 0);
            NPOI.HSSF.Record.Record[] outRecs = rv.Records;

            Assert.AreEqual(recs.Length, outRecs.Length);
            //expected order of records:
            NPOI.HSSF.Record.Record[] expectedRecs = {
                recs[0],  //BOFRecord
                recs[1],  //IndexRecord

                //PageSettingsBlock
                recs[2],  //HeaderRecord
                recs[3],  //FooterRecord
                recs[9],  //HeaderFooterRecord
                recs[4],  // DimensionsRecord
                recs[5],  // WindowTwoRecord

                //CustomViewSettingsRecordAggregate
                recs[6],  // UserSViewBegin
                recs[7],  // SelectionRecord
                recs[10], // HeaderFooterRecord
                recs[8],  // UserSViewEnd

                recs[11],  //EOFRecord
            };
            for (int i = 0; i < expectedRecs.Length; i++)
            {
                Assert.AreEqual(expectedRecs[i].GetType(), outRecs[i].GetType(), "Record mismatch at index " + i);
            }
            HeaderFooterRecord hd1 = (HeaderFooterRecord)expectedRecs[4];
            //GUID is zero
            Assert.IsTrue(Arrays.Equals(new byte[16], hd1.Guid));
            Assert.IsTrue(hd1.IsCurrentSheet);

            UserSViewBegin svb = (UserSViewBegin)expectedRecs[7];
            HeaderFooterRecord hd2 = (HeaderFooterRecord)expectedRecs[9];
            Assert.IsFalse(hd2.IsCurrentSheet);
            //GUIDs of HeaderFooterRecord and UserSViewBegin must be the same
            Assert.IsTrue(Arrays.Equals(svb.Guid, hd2.Guid));
        }

    }
}

