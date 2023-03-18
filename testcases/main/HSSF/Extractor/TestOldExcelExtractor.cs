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
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using System.Text;

    /**
     * Unit tests for the Excel 5/95 and Excel 4 (and older) text 
     *  extractor
     */
    [TestFixture]
    public class TestOldExcelExtractor : POITestCase
    {
        private static OldExcelExtractor CreateExtractor(String sampleFileName)
        {
            FileInfo file = HSSFTestDataSamples.GetSampleFile(sampleFileName);

            return new OldExcelExtractor(file);
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

            extractor.Close();
        }
        [Test]
        public void TestSimpleExcel3NoReading()
        {
            OldExcelExtractor extractor = CreateExtractor("testEXCEL_3.xls");
            Assert.IsNotNull(extractor);

            extractor.Close();
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

            extractor.Close();
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

                extractor.Close();
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

            extractor.Close();
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

            extractor.Close();
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

                extractor.Close();
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

                extractor.Close();
            }
        }

        [Test]//(expected=OfficeXmlFileException.class)
        public void TestOpenInvalidFile1() => Assert.Throws<OfficeXmlFileException>(() =>
        {
            // a file that exists, but is a different format
            CreateExtractor("WithVariousData.xlsx");
        });

        [Test]//(expected=RecordFormatException.class)
        public void TestOpenInvalidFile2() => Assert.Throws<RecordFormatException>(() =>
        {
            // a completely different type of file
            CreateExtractor("48936-strings.txt");
        });

        [Test]//(expected=FileNotFoundException.class)
        public void TestOpenInvalidFile3() => Assert.Throws<NotSupportedException>(() =>
        {
            // a POIFS file which is not a Workbook
            Stream @is = POIDataSamples.GetDocumentInstance().OpenResourceAsStream("47304.doc");
            try
            {
                new OldExcelExtractor(@is).Close();
            }
            finally
            {
                @is.Close();
            }
        });

        [Test]//(expected= EmptyFileException.class)
        public void TestOpenNonExistingFile() => Assert.Throws<FileNotFoundException>(() =>
        {
            // a file that exists, but is a different format
            OldExcelExtractor extractor = new OldExcelExtractor(new FileInfo("notexistingfile.xls"));
            extractor.Close();
        });

        [Test]
        public void TestInputStream()
        {
            FileInfo file = HSSFTestDataSamples.GetSampleFile("testEXCEL_3.xls");
            Stream stream = file.OpenRead();
            try
            {
                OldExcelExtractor extractor = new OldExcelExtractor(stream);
                String text = extractor.Text;
                Assert.IsNotNull(text);
                extractor.Close();
            }
            finally
            {
                stream.Close();
            }
        }
        [Test]
        public void TestInputStreamNPOIHeader()
        {
            FileInfo file = HSSFTestDataSamples.GetSampleFile("FormulaRefs.xls");
            Stream stream = file.OpenRead();
            try
            {
                OldExcelExtractor extractor = new OldExcelExtractor(stream);
                extractor.Close();
            }
            finally
            {
                stream.Close();
            }
        }
        [Test]
        public void TestNPOIFSFileSystem()
        {
            FileInfo file = HSSFTestDataSamples.GetSampleFile("FormulaRefs.xls");
            NPOIFSFileSystem fs = new NPOIFSFileSystem(file);
            try
            {
                OldExcelExtractor extractor = new OldExcelExtractor(fs);
                extractor.Close();
            }
            finally
            {
                fs.Close();
            }
        }
        [Test]
        public void TestDirectoryNode()
        {
            FileInfo file = HSSFTestDataSamples.GetSampleFile("FormulaRefs.xls");
            NPOIFSFileSystem fs = new NPOIFSFileSystem(file);
            try
            {
                OldExcelExtractor extractor = new OldExcelExtractor(fs.Root);
                extractor.Close();
            }
            finally
            {
                fs.Close();
            }
        }
        [Test]
        public void TestDirectoryNodeInvalidFile()
        {
            FileStream file = POIDataSamples.GetDocumentInstance().GetFile("test.doc");
            NPOIFSFileSystem fs = new NPOIFSFileSystem(file);
            try
            {
                OldExcelExtractor extractor = new OldExcelExtractor(fs.Root);
                extractor.Close();
                Assert.Fail("Should catch exception here");
            }
            catch (FileNotFoundException)
            {
                // expected here
            }
            finally
            {
                fs.Close();
            }
        }
        [Ignore("Calls System.exit()")]
        [Test]
        public void TestMainUsage()
        {
            TextWriter save = System.Console.Error;
            try
            {
                ByteArrayOutputStream out1 = new ByteArrayOutputStream();
                try
                {

                    TextWriter str = new StreamWriter(out1, Encoding.UTF8) { AutoFlush = false };
                    Console.SetError(str);
                    OldExcelExtractor.main(new String[] { });
                }
                finally
                {
                out1.Close();
                }
            }
            finally
            {
                Console.SetError(save);
            }
        }
        [Test]
        public void TestMain()
        {
            FileInfo file = HSSFTestDataSamples.GetSampleFile("testEXCEL_3.xls");
            TextWriter save = Console.Out;
            try
            {
                ByteArrayOutputStream out1 = new ByteArrayOutputStream();
                try
                {
                    TextWriter str = new StreamWriter(out1, Encoding.UTF8) { AutoFlush = false };
                    Console.SetOut(str);
                    OldExcelExtractor.main(new String[] { file.FullName });
                }
                finally
                {
                out1.Close();
                }
                String string1 = Encoding.UTF8.GetString(out1.ToByteArray());
                Assert.IsTrue(string1.Contains("Table C-13--Lemons"), "Had: " + string1);
            }
            finally
            {
                Console.SetOut(save);
            }
        }

    }

}
