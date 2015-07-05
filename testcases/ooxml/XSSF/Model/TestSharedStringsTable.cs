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

using NUnit.Framework;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System;
using NPOI.SS.UserModel;
using TestCases;
using System.IO;
namespace NPOI.XSSF.Model
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
            Assert.IsNotNull(sst.Items);
            Assert.AreEqual(0, sst.Items.Count);
            Assert.AreEqual(0, sst.Count);
            Assert.AreEqual(0, sst.UniqueCount);

            st = new CT_Rst();
            st.t = ("Hello, World!");

            idx = sst.AddEntry(st);
            Assert.AreEqual(0, idx);
            Assert.AreEqual(1, sst.Count);
            Assert.AreEqual(1, sst.UniqueCount);

            //add the same entry again
            idx = sst.AddEntry(st);
            Assert.AreEqual(0, idx);
            Assert.AreEqual(2, sst.Count);
            Assert.AreEqual(1, sst.UniqueCount);

            //and again
            idx = sst.AddEntry(st);
            Assert.AreEqual(0, idx);
            Assert.AreEqual(3, sst.Count);
            Assert.AreEqual(1, sst.UniqueCount);

            st = new CT_Rst();
            st.t = ("Second string");

            idx = sst.AddEntry(st);
            Assert.AreEqual(1, idx);
            Assert.AreEqual(4, sst.Count);
            Assert.AreEqual(2, sst.UniqueCount);

            //add the same entry again
            idx = sst.AddEntry(st);
            Assert.AreEqual(1, idx);
            Assert.AreEqual(5, sst.Count);
            Assert.AreEqual(2, sst.UniqueCount);

            st = new CT_Rst();
            CT_RElt r = st.AddNewR();
            CT_RPrElt pr = r.AddNewRPr();
            pr.AddNewColor().SetRgb(new byte[] { (byte)0xFF, 0, 0 }); //red
            pr.AddNewI().val = (true);  //bold
            pr.AddNewB().val = (true);  //italic
            r.t = ("Second string");

            idx = sst.AddEntry(st);
            Assert.AreEqual(2, idx);
            Assert.AreEqual(6, sst.Count);
            Assert.AreEqual(3, sst.UniqueCount);

            idx = sst.AddEntry(st);
            Assert.AreEqual(2, idx);
            Assert.AreEqual(7, sst.Count);
            Assert.AreEqual(3, sst.UniqueCount);

            //OK. the sst table is Filled, check the contents
            Assert.AreEqual(3, sst.Items.Count);
            Assert.AreEqual("Hello, World!", new XSSFRichTextString(sst.GetEntryAt(0)).ToString());
            Assert.AreEqual("Second string", new XSSFRichTextString(sst.GetEntryAt(1)).ToString());
            Assert.AreEqual("Second string", new XSSFRichTextString(sst.GetEntryAt(2)).ToString());
        }
        [Test]
        public void TestReadWrite()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("sample.xlsx");
            SharedStringsTable sst1 = wb.GetSharedStringSource();

            //Serialize, read back and compare with the original
            SharedStringsTable sst2 = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb)).GetSharedStringSource();

            Assert.AreEqual(sst1.Count, sst2.Count);
            Assert.AreEqual(sst1.UniqueCount, sst2.UniqueCount);

            List<CT_Rst> items1 = sst1.Items;
            List<CT_Rst> items2 = sst2.Items;
            Assert.AreEqual(items1.Count, items2.Count);
            for (int i = 0; i < items1.Count; i++)
            {
                CT_Rst st1 = items1[i];
                CT_Rst st2 = items2[i];
                Assert.AreEqual(st1.ToString(), st2.ToString());
            }

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
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
                Assert.AreEqual(str, val);
            }

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(w));
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

    }
}

