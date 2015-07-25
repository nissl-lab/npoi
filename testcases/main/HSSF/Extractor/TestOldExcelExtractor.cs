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

namespace TestCases.HSSF.Extractor
{
    using System;
    using NPOI.HSSF;
    using TestCases;
    using NUnit.Framework;
    using System.IO;
    using TestCases.HSSF;
    using NPOI.HSSF.Extractor;

    /**
     * Unit tests for the Excel 5/95 and Excel 4 (and older) text 
     *  extractor
     */
    [TestFixture]
    public class TestOldExcelExtractor : POITestCase
    {
        private static OldExcelExtractor CreateExtractor(String sampleFileName)
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);

            try
            {
                MemoryStream ms = new MemoryStream();
                is1.CopyTo(ms);
                ms.Position = 0;
                return new OldExcelExtractor(ms);


            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                //is1.Close();
            }
        }

        [Test]
        public void TestSimpleExcel3()
        {
            OldExcelExtractor extractor = CreateExtractor("testEXCEL_3.xls");

            // Check we can call GetText without error
            String text = extractor.Text;

            // Check we find a few words we expect in there
            AssertContains(text, "Season beginning August");
            AssertContains(text, "USDA");

            // Check we find a few numbers we expect in there
            AssertContains(text, "347");
            AssertContains(text, "228");

            // Check we find a few string-literal dates in there
            AssertContains(text, "1981/82");

            // Check the type
            Assert.AreEqual(3, extractor.BiffVersion);
            Assert.AreEqual(0x10, extractor.FileType);
        }
        [Test]
        public void TestSimpleExcel4()
        {
            OldExcelExtractor extractor = CreateExtractor("testEXCEL_4.xls");

            // Check we can call GetText without error
            String text = extractor.Text;

            // Check we find a few words we expect in there
            AssertContains(text, "Size");
            AssertContains(text, "Returns");

            // Check we find a few numbers we expect in there
            AssertContains(text, "11");
            AssertContains(text, "784");

            // Check the type
            Assert.AreEqual(4, extractor.BiffVersion);
            Assert.AreEqual(0x10, extractor.FileType);
        }
        [Test]
        public void TestSimpleExcel5()
        {
            foreach (String ver in new String[] { "5", "95" })
            {
                OldExcelExtractor extractor = CreateExtractor("testEXCEL_" + ver + ".xls");

                // Check we can call GetText without error
                String text = extractor.Text;

                // Check we find a few words we expect in there
                AssertContains(text, "Sample Excel");
                AssertContains(text, "Written and saved");

                // Check we find a few numbers we expect in there
                AssertContains(text, "15");
                AssertContains(text, "169");

                // Check we got the sheet names (new formats only)
                AssertContains(text, "Sheet: Feuil3");

                // Check the type
                Assert.AreEqual(5, extractor.BiffVersion);
                Assert.AreEqual(0x05, extractor.FileType);
            }
        }

        [Test]
        public void TestStrings()
        {
            OldExcelExtractor extractor = CreateExtractor("testEXCEL_4.xls");
            String text = extractor.Text;

            // Simple strings
            AssertContains(text, "Table 10 -- Examination Coverage:");
            AssertContains(text, "Recommended and Average Recommended Additional Tax After");
            AssertContains(text, "Individual income tax returns, total");

            // More complicated strings
            AssertContains(text, "$100,000 or more");
            AssertContains(text, "S corporation returns, Form 1120S [10,15]");
            AssertContains(text, "individual income tax return \u201Cshort forms.\u201D");

            // Formula based strings
            // TODO Find some then test
        }

        [Test]
        public void TestFormattedNumbersExcel4()
        {
            OldExcelExtractor extractor = CreateExtractor("testEXCEL_4.xls");
            String text = extractor.Text;

            // Simple numbers
            AssertContains(text, "151");
            AssertContains(text, "784");

            // Numbers which come from formulas
            AssertContains(text, "0.398"); // TODO Rounding
            AssertContains(text, "624");

            // Formatted numbers
            // TODO
            //      AssertContains(text, "55,624");
            //      AssertContains(text, "11,743,477");
        }
        [Test]
        public void TestFormattedNumbersExcel5()
        {
            foreach (String ver in new String[] { "5", "95" })
            {
                OldExcelExtractor extractor = CreateExtractor("testEXCEL_" + ver + ".xls");
                String text = extractor.Text;

                // Simple numbers
                AssertContains(text, "1");

                // Numbers which come from formulas
                AssertContains(text, "13");
                AssertContains(text, "169");

                // Formatted numbers
                // TODO
                //          AssertContains(text, "100.00%");
                //          AssertContains(text, "155.00%");
                //          AssertContains(text, "1,125");
                //          AssertContains(text, "189,945");
                //          AssertContains(text, "1,234,500");
                //          AssertContains(text, "$169.00");
                //          AssertContains(text, "$1,253.82");
            }
        }

        [Test]
        public void TestFromFile()
        {
            foreach (String ver in new String[] { "4", "5", "95" })
            {
                String filename = "testEXCEL_" + ver + ".xls";
                FileInfo f = HSSFTestDataSamples.GetSampleFile(filename);

                OldExcelExtractor extractor = new OldExcelExtractor(f);
                String text = extractor.Text;
                Assert.IsNotNull(text);
                Assert.IsTrue(text.Length > 100);
            }
        }
    }

}