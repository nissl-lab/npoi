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

using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System;
using NPOI.SS.UserModel;
using TestCases;
using System.IO;
using NPOI.XSSF.Model;
using NPOI.XSSF;
using NPOI;

namespace TestCases.XSSF.Model
{

    /**
     * Test {@link SharedStringsTable}, the cache of strings in a workbook
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestSharedStringsTable
    {
        [Test]
        public void TestCreateNew()
        {
            SharedStringsTable sst = new SharedStringsTable();

            CT_Rst st;
            int idx;

            // Check defaults
            ClassicAssert.IsNotNull(sst.Items);
            ClassicAssert.AreEqual(0, sst.Items.Count);
            ClassicAssert.AreEqual(0, sst.Count);
            ClassicAssert.AreEqual(0, sst.UniqueCount);

            st = new CT_Rst();
            st.t = ("Hello, World!");

            idx = sst.AddEntry(st);
            ClassicAssert.AreEqual(0, idx);
            ClassicAssert.AreEqual(1, sst.Count);
            ClassicAssert.AreEqual(1, sst.UniqueCount);

            //add the same entry again
            idx = sst.AddEntry(st);
            ClassicAssert.AreEqual(0, idx);
            ClassicAssert.AreEqual(2, sst.Count);
            ClassicAssert.AreEqual(1, sst.UniqueCount);

            //and again
            idx = sst.AddEntry(st);
            ClassicAssert.AreEqual(0, idx);
            ClassicAssert.AreEqual(3, sst.Count);
            ClassicAssert.AreEqual(1, sst.UniqueCount);

            st = new CT_Rst();
            st.t = ("Second string");

            idx = sst.AddEntry(st);
            ClassicAssert.AreEqual(1, idx);
            ClassicAssert.AreEqual(4, sst.Count);
            ClassicAssert.AreEqual(2, sst.UniqueCount);

            //add the same entry again
            idx = sst.AddEntry(st);
            ClassicAssert.AreEqual(1, idx);
            ClassicAssert.AreEqual(5, sst.Count);
            ClassicAssert.AreEqual(2, sst.UniqueCount);

            st = new CT_Rst();
            CT_RElt r = st.AddNewR();
            CT_RPrElt pr = r.AddNewRPr();
            pr.AddNewColor().SetRgb(new byte[] { (byte)0xFF, 0, 0 }); //red
            pr.AddNewI().val = (true);  //bold
            pr.AddNewB().val = (true);  //italic
            r.t = ("Second string");

            idx = sst.AddEntry(st);
            ClassicAssert.AreEqual(2, idx);
            ClassicAssert.AreEqual(6, sst.Count);
            ClassicAssert.AreEqual(3, sst.UniqueCount);

            idx = sst.AddEntry(st);
            ClassicAssert.AreEqual(2, idx);
            ClassicAssert.AreEqual(7, sst.Count);
            ClassicAssert.AreEqual(3, sst.UniqueCount);

            //OK. the sst table is Filled, check the contents
            ClassicAssert.AreEqual(3, sst.Items.Count);
            ClassicAssert.AreEqual("Hello, World!", new XSSFRichTextString(sst.GetEntryAt(0)).ToString());
            ClassicAssert.AreEqual("Second string", new XSSFRichTextString(sst.GetEntryAt(1)).ToString());
            ClassicAssert.AreEqual("Second string", new XSSFRichTextString(sst.GetEntryAt(2)).ToString());
        }
        [Test]
        public void TestReadWrite()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            SharedStringsTable sst1 = wb.GetSharedStringSource();

            //Serialize, read back and compare with the original
            SharedStringsTable sst2 = XSSFTestDataSamples.WriteOutAndReadBack(wb).GetSharedStringSource();

            ClassicAssert.AreEqual(sst1.Count, sst2.Count);
            ClassicAssert.AreEqual(sst1.UniqueCount, sst2.UniqueCount);

            IList<CT_Rst> items1 = sst1.Items;
            IList<CT_Rst> items2 = sst2.Items;
            ClassicAssert.AreEqual(items1.Count, items2.Count);
            for (int i = 0; i < items1.Count; i++)
            {
                CT_Rst st1 = items1[i];
                CT_Rst st2 = items2[i];
                ClassicAssert.AreEqual(st1.ToString(), st2.ToString());
            }

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }

        /**
         * Test for Bugzilla 48936
         *
         * A specific sequence of strings can result in broken CDATA section in sharedStrings.xml file.
         *
         * @author Philippe Laflamme
         */
        [Test]
        public void TestBug48936()
        {
            IWorkbook w = new XSSFWorkbook();
            ISheet s = w.CreateSheet();
            int i = 0;
            List<String> lst = ReadStrings("48936-strings.txt");
            foreach (String str in lst)
            {
                s.CreateRow(i++).CreateCell(0).SetCellValue(str);
            }

            try
            {
                w = XSSFTestDataSamples.WriteOutAndReadBack(w);
            }
            catch (POIXMLException)
            {
                Assert.Fail("Detected Bug #48936");
            }
            s = w.GetSheetAt(0);
            i = 0;
            foreach (String str in lst)
            {
                String val = s.GetRow(i++).GetCell(0).StringCellValue;
                ClassicAssert.AreEqual(str, val);
            }

            ClassicAssert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(w));
        }

        private List<String> ReadStrings(String filename)
        {
            List<String> strs = new List<String>();
            POIDataSamples samples = POIDataSamples.GetSpreadSheetInstance();
            StreamReader br =
                    new StreamReader(samples.OpenResourceAsStream(filename));
            String s;
            while ((s = br.ReadLine()) != null)
            {
                if (s.Trim().Length > 0)
                {
                    strs.Add(s.Trim());
                }
            }
            br.Close();
            return strs;
        }

        /// <summary>
        /// Verify that opening a workbook and writing it without accessing shared
        /// strings does not cause the SST to be parsed or rewritten.
        /// </summary>
        [Test]
        public void TestLazyLoadNotTriggeredByWrite()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            SharedStringsTable sst = wb.GetSharedStringSource();

            // SST should NOT be loaded yet
            ClassicAssert.IsFalse(sst.IsLoaded, "SST should not be loaded before any access");

            // Write the workbook without touching any string cells
            byte[] writtenBytes;
            using (MemoryStream ms = new MemoryStream())
            {
                wb.Write(ms, false);
                writtenBytes = ms.ToArray();
            }

            // SST should still be unloaded (not dirty, not accessed)
            ClassicAssert.IsFalse(sst.IsLoaded, "SST should still not be loaded after Write() without string access");

            // The written workbook should still contain the correct SST
            XSSFWorkbook wb2 = new XSSFWorkbook(new MemoryStream(writtenBytes));
            SharedStringsTable sst2 = wb2.GetSharedStringSource();

            // Now access to force load
            ClassicAssert.IsTrue(sst2.Count > 0, "Written workbook should have preserved the SST");
            ClassicAssert.IsTrue(sst2.Items.Count > 0);

            wb.Close();
            wb2.Close();
        }

        /// <summary>
        /// Verify that reading SST content marks it loaded but not dirty,
        /// and that saving preserves the original SST data without re-serializing.
        /// </summary>
        [Test]
        public void TestReadSstNotDirtyAfterAccess()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            SharedStringsTable sst1 = wb1.GetSharedStringSource();

            // Access SST – this should load but not dirty it
            int origCount = sst1.Count;
            int origUnique = sst1.UniqueCount;
            IList<CT_Rst> origItems = sst1.Items;
            ClassicAssert.IsTrue(sst1.IsLoaded, "SST should be loaded after Count access");

            // Round-trip: write + read back
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1);
            SharedStringsTable sst2 = wb2.GetSharedStringSource();

            ClassicAssert.AreEqual(origCount, sst2.Count);
            ClassicAssert.AreEqual(origUnique, sst2.UniqueCount);
            ClassicAssert.AreEqual(origItems.Count, sst2.Items.Count);
            for (int i = 0; i < origItems.Count; i++)
                ClassicAssert.AreEqual(origItems[i].ToString(), sst2.Items[i].ToString());

            wb1.Close();
            wb2.Close();
        }

        /// <summary>
        /// Verify that rich text runs and phonetic runs in 51519.xlsx are parsed
        /// correctly by the streaming parser and survive a round-trip write + read.
        /// </summary>
        [Test]
        public void TestPhoneticAndRichTextFidelity()
        {
            POIDataSamples ssTests = POIDataSamples.GetSpreadSheetInstance();
            XSSFWorkbook wb = new XSSFWorkbook(ssTests.OpenResourceAsStream("51519.xlsx"));
            SharedStringsTable sst = wb.GetSharedStringSource();

            ClassicAssert.AreEqual(49, sst.Items.Count, "Expected 49 shared strings in 51519.xlsx");

            // Entry 0: plain Japanese text (no rich runs)
            CT_Rst entry0 = sst.GetEntryAt(0);
            ClassicAssert.AreEqual("\u30B3\u30E1\u30F3\u30C8",
                new XSSFRichTextString(entry0).ToString(),
                "Entry 0 text mismatch");

            // Entry 3: should have phonetic runs (rPh elements)
            CT_Rst entry3 = sst.GetEntryAt(3);
            ClassicAssert.IsNotNull(entry3.rPh, "Entry 3 should have phonetic runs");
            ClassicAssert.IsTrue(entry3.rPh.Count > 0, "Entry 3 should have at least one phonetic run");

            // Round-trip: write + read back
            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            SharedStringsTable sst2 = wb2.GetSharedStringSource();

            ClassicAssert.AreEqual(49, sst2.Items.Count, "Round-tripped SST should still have 49 entries");

            CT_Rst entry0rt = sst2.GetEntryAt(0);
            ClassicAssert.AreEqual("\u30B3\u30E1\u30F3\u30C8",
                new XSSFRichTextString(entry0rt).ToString(),
                "Entry 0 text mismatch after round-trip");

            CT_Rst entry3rt = sst2.GetEntryAt(3);
            ClassicAssert.IsNotNull(entry3rt.rPh, "Entry 3 should have phonetic runs after round-trip");
            ClassicAssert.AreEqual(entry3.rPh.Count, entry3rt.rPh.Count,
                "Phonetic run count should match after round-trip");

            wb.Close();
            wb2.Close();
        }

    }
}

