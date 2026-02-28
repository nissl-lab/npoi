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

using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NPOI.HWPF.Model;

using System;
using NPOI;
using NPOI.HWPF.Extractor;
using NUnit.Framework;
namespace TestCases.HWPF.UserModel
{

    /**
     * Test various problem documents
     *
     * @author Nick Burch (nick at torchbox dot com)
     */
    [TestFixture]
    public class TestProblems : HWPFTestCase
    {

        /**
         * ListEntry passed no ListTable
         */
        [Test]
        public void TestListEntryNoListTable()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("ListEntryNoListTable.doc");

            Range r = doc.GetRange();
            StyleSheet styleSheet = doc.GetStyleSheet();
            for (int x = 0; x < r.NumSections; x++)
            {
                Section s = r.GetSection(x);
                for (int y = 0; y < s.NumParagraphs; y++)
                {
                    Paragraph paragraph = s.GetParagraph(y);
                    // Console.WriteLine(paragraph.GetCharacterRun(0).Text);
                }
            }
        }

        /**
         * AIOOB for TableSprmUncompressor.unCompressTAPOperation
         */
        [Test]
        public void TestSprmAIOOB()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("AIOOB-Tap.doc");

            Range r = doc.GetRange();
            StyleSheet styleSheet = doc.GetStyleSheet();
            for (int x = 0; x < r.NumSections; x++)
            {
                Section s = r.GetSection(x);
                for (int y = 0; y < s.NumParagraphs; y++)
                {
                    Paragraph paragraph = s.GetParagraph(y);
                    // Console.WriteLine(paragraph.GetCharacterRun(0).Text);
                }
            }
        }

        /**
         * Test for TableCell not skipping the last paragraph. Bugs #45062 and
         * #44292
         */
        [Test]
        public void TestTableCellLastParagraph()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug44292.doc");
            Range r = doc.GetRange();
            Assert.AreEqual(6, r.NumParagraphs);
            Assert.AreEqual(0, r.StartOffset);
            Assert.AreEqual(87, r.EndOffset);

            // Paragraph with table
            Paragraph p = r.GetParagraph(0);
            Assert.AreEqual(0, p.StartOffset);
            Assert.AreEqual(20, p.EndOffset);

            // Check a few bits of the table directly
            Assert.AreEqual("One paragraph is ok\u0007", r.GetParagraph(0).Text);
            Assert.AreEqual("First para is ok\r", r.GetParagraph(1).Text);
            Assert.AreEqual("Second paragraph is skipped\u0007", r.GetParagraph(2).Text);
            Assert.AreEqual("One paragraph is ok\u0007", r.GetParagraph(3).Text);
            Assert.AreEqual("\u0007", r.GetParagraph(4).Text);
            Assert.AreEqual("\r", r.GetParagraph(5).Text);

            // Get the table
            Table t = r.GetTable(p);

            // get the only row
            Assert.AreEqual(1, t.NumRows);
            TableRow row = t.GetRow(0);

            // sanity check our row
            Assert.AreEqual(5, row.NumParagraphs);
            Assert.AreEqual(0, row._parStart);
            Assert.AreEqual(5, row._parEnd);
            Assert.AreEqual(0, row.StartOffset);
            Assert.AreEqual(86, row.EndOffset);


            // get the first cell
            TableCell cell = row.GetCell(0);
            // First cell should have one paragraph
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual("One paragraph is ok\u0007", cell.GetParagraph(0).Text);
            Assert.AreEqual(0, cell._parStart);
            Assert.AreEqual(1, cell._parEnd);
            Assert.AreEqual(0, cell.StartOffset);
            Assert.AreEqual(20, cell.EndOffset);


            // get the second
            cell = row.GetCell(1);
            // Second cell should be detected as having two paragraphs
            Assert.AreEqual(2, cell.NumParagraphs);
            Assert.AreEqual("First para is ok\r", cell.GetParagraph(0).Text);
            Assert.AreEqual("Second paragraph is skipped\u0007", cell.GetParagraph(1).Text);
            Assert.AreEqual(1, cell._parStart);
            Assert.AreEqual(3, cell._parEnd);
            Assert.AreEqual(20, cell.StartOffset);
            Assert.AreEqual(65, cell.EndOffset);


            // get the last cell
            cell = row.GetCell(2);
            // Last cell should have one paragraph
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual("One paragraph is ok\u0007", cell.GetParagraph(0).Text);
            Assert.AreEqual(3, cell._parStart);
            Assert.AreEqual(4, cell._parEnd);
            Assert.AreEqual(65, cell.StartOffset);
            Assert.AreEqual(85, cell.EndOffset);
        }
        [Test]
        public void TestRangeDelete()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug28627.doc");

            Range range = doc.GetRange();
            int numParagraphs = range.NumParagraphs;

            int totalLength = 0, deletedLength = 0;

            for (int i = 0; i < numParagraphs; i++)
            {
                Paragraph para = range.GetParagraph(i);
                String text = para.Text;

                totalLength += text.Length;
                if (text.IndexOf("{delete me}") > -1)
                {
                    para.Delete();
                    deletedLength += text.Length;
                }
            }

            // check the text length after deletion
            int newLength = 0;
            range = doc.GetRange();
            numParagraphs = range.NumParagraphs;

            for (int i = 0; i < numParagraphs; i++)
            {
                Paragraph para = range.GetParagraph(i);
                String text = para.Text;

                newLength += text.Length;
            }

            Assert.AreEqual(newLength, totalLength - deletedLength);
        }

        /**
         * With an encrypted file, we should give a suitable exception, and not OOM
         */
        [Test]
        public void TestEncryptedFile()
        {
            try
            {
                HWPFTestDataSamples.OpenSampleFile("PasswordProtected.doc");
                Assert.Fail();
            }
            catch (EncryptedDocumentException )
            {
                // Good
            }
        }
        [Test]
        public void TestWriteProperties()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("SampleDoc.doc");
            Assert.AreEqual("Nick Burch", doc.SummaryInformation.Author);

            // Write and read
            HWPFDocument doc2 = WriteOutAndRead(doc);
            Assert.AreEqual("Nick Burch", doc2.SummaryInformation.Author);
        }

        /**
         * Test for Reading paragraphs from Range after replacing some 
         * text in this Range.
         * Bug #45269
         */
        [Test]
        public void TestReadParagraphsAfterReplaceText()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug45269.doc");
            Range range = doc.GetRange();

            String toFind = "campo1";
            String longer = " foi porraaaaa ";
            String shorter = " foi ";

            //check replace with longer text
            for (int x = 0; x < range.NumParagraphs; x++)
            {
                Paragraph para = range.GetParagraph(x);
                int offset = para.Text.IndexOf(toFind);
                if (offset >= 0)
                {
                    para.ReplaceText(toFind, longer, offset);
                    Assert.AreEqual(offset, para.Text.IndexOf(longer));
                }
            }

            doc = HWPFTestDataSamples.OpenSampleFile("Bug45269.doc");
            range = doc.GetRange();

            //check replace with shorter text
            for (int x = 0; x < range.NumParagraphs; x++)
            {
                Paragraph para = range.GetParagraph(x);
                int offset = para.Text.IndexOf(toFind);
                if (offset >= 0)
                {
                    para.ReplaceText(toFind, shorter, offset);
                    Assert.AreEqual(offset, para.Text.IndexOf(shorter));
                }
            }
        }

        /**
         * Bug #49936 - Problems with Reading the header out of
         *  the Header Stories
         */

        [Test]
        public void TestProblemHeaderStories49936()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("HeaderFooterProblematic.doc");
            HeaderStories hs = new HeaderStories(doc);

            Assert.AreEqual("", hs.FirstHeader);
            Assert.AreEqual("\r", hs.EvenHeader);
            Assert.AreEqual("", hs.OddHeader);

            Assert.AreEqual("", hs.FirstFooter);
            Assert.AreEqual("", hs.EvenFooter);
            Assert.AreEqual("", hs.OddFooter);

            WordExtractor ext = new WordExtractor(doc);
            Assert.AreEqual("\n", ext.HeaderText);
            Assert.AreEqual("", ext.FooterText);
        }

        /**
         * Bug #45877 - problematic PAPX with no parent set
         */
        [Test]
        public void TestParagraphPAPXNoParent45877()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug45877.doc");
            Assert.AreEqual(17, doc.GetRange().NumParagraphs);

            Assert.AreEqual("First paragraph\r", doc.GetRange().GetParagraph(0).Text);
            Assert.AreEqual("After Crashing Part\r", doc.GetRange().GetParagraph(13).Text);
        }

        /**
         * Bug #48245 - don't include the text from the
         *  next cell in the current one
         */
        [Test]
        public void TestTableIterator()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("simple-table2.doc");
            Range r = doc.GetRange();

            // Check the text is as we'd expect
            Assert.AreEqual(13, r.NumParagraphs);
            Assert.AreEqual("Row 1/Cell 1\u0007", r.GetParagraph(0).Text);
            Assert.AreEqual("Row 1/Cell 2\u0007", r.GetParagraph(1).Text);
            Assert.AreEqual("Row 1/Cell 3\u0007", r.GetParagraph(2).Text);
            Assert.AreEqual("\u0007", r.GetParagraph(3).Text);
            Assert.AreEqual("Row 2/Cell 1\u0007", r.GetParagraph(4).Text);
            Assert.AreEqual("Row 2/Cell 2\u0007", r.GetParagraph(5).Text);
            Assert.AreEqual("Row 2/Cell 3\u0007", r.GetParagraph(6).Text);
            Assert.AreEqual("\u0007", r.GetParagraph(7).Text);
            Assert.AreEqual("Row 3/Cell 1\u0007", r.GetParagraph(8).Text);
            Assert.AreEqual("Row 3/Cell 2\u0007", r.GetParagraph(9).Text);
            Assert.AreEqual("Row 3/Cell 3\u0007", r.GetParagraph(10).Text);
            Assert.AreEqual("\u0007", r.GetParagraph(11).Text);
            Assert.AreEqual("\r", r.GetParagraph(12).Text);

            Paragraph p;

            // Take a look in detail at the first couple of
            //  paragraphs
            p = r.GetParagraph(0);
            Assert.AreEqual(1, p.NumParagraphs);
            Assert.AreEqual(0, p.StartOffset);
            Assert.AreEqual(13, p.EndOffset);
            Assert.AreEqual(0, p._parStart);
            Assert.AreEqual(1, p._parEnd);

            p = r.GetParagraph(1);
            Assert.AreEqual(1, p.NumParagraphs);
            Assert.AreEqual(13, p.StartOffset);
            Assert.AreEqual(26, p.EndOffset);
            Assert.AreEqual(1, p._parStart);
            Assert.AreEqual(2, p._parEnd);

            p = r.GetParagraph(2);
            Assert.AreEqual(1, p.NumParagraphs);
            Assert.AreEqual(26, p.StartOffset);
            Assert.AreEqual(39, p.EndOffset);
            Assert.AreEqual(2, p._parStart);
            Assert.AreEqual(3, p._parEnd);


            // Now look at the table
            Table table = r.GetTable(r.GetParagraph(0));
            Assert.AreEqual(3, table.NumRows);

            TableRow row;
            TableCell cell;


            row = table.GetRow(0);
            Assert.AreEqual(0, row._parStart);
            Assert.AreEqual(4, row._parEnd);

            cell = row.GetCell(0);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(0, cell._parStart);
            Assert.AreEqual(1, cell._parEnd);
            Assert.AreEqual(0, cell.StartOffset);
            Assert.AreEqual(13, cell.EndOffset);
            Assert.AreEqual("Row 1/Cell 1\u0007", cell.Text);
            Assert.AreEqual("Row 1/Cell 1\u0007", cell.GetParagraph(0).Text);

            cell = row.GetCell(1);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(1, cell._parStart);
            Assert.AreEqual(2, cell._parEnd);
            Assert.AreEqual(13, cell.StartOffset);
            Assert.AreEqual(26, cell.EndOffset);
            Assert.AreEqual("Row 1/Cell 2\u0007", cell.Text);
            Assert.AreEqual("Row 1/Cell 2\u0007", cell.GetParagraph(0).Text);

            cell = row.GetCell(2);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(2, cell._parStart);
            Assert.AreEqual(3, cell._parEnd);
            Assert.AreEqual(26, cell.StartOffset);
            Assert.AreEqual(39, cell.EndOffset);
            Assert.AreEqual("Row 1/Cell 3\u0007", cell.Text);
            Assert.AreEqual("Row 1/Cell 3\u0007", cell.GetParagraph(0).Text);


            // Onto row #2
            row = table.GetRow(1);
            Assert.AreEqual(4, row._parStart);
            Assert.AreEqual(8, row._parEnd);

            cell = row.GetCell(0);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(4, cell._parStart);
            Assert.AreEqual(5, cell._parEnd);
            Assert.AreEqual(40, cell.StartOffset);
            Assert.AreEqual(53, cell.EndOffset);
            Assert.AreEqual("Row 2/Cell 1\u0007", cell.Text);

            cell = row.GetCell(1);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(5, cell._parStart);
            Assert.AreEqual(6, cell._parEnd);
            Assert.AreEqual(53, cell.StartOffset);
            Assert.AreEqual(66, cell.EndOffset);
            Assert.AreEqual("Row 2/Cell 2\u0007", cell.Text);

            cell = row.GetCell(2);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(6, cell._parStart);
            Assert.AreEqual(7, cell._parEnd);
            Assert.AreEqual(66, cell.StartOffset);
            Assert.AreEqual(79, cell.EndOffset);
            Assert.AreEqual("Row 2/Cell 3\u0007", cell.Text);


            // Finally row 3
            row = table.GetRow(2);
            Assert.AreEqual(8, row._parStart);
            Assert.AreEqual(12, row._parEnd);

            cell = row.GetCell(0);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(8, cell._parStart);
            Assert.AreEqual(9, cell._parEnd);
            Assert.AreEqual(80, cell.StartOffset);
            Assert.AreEqual(93, cell.EndOffset);
            Assert.AreEqual("Row 3/Cell 1\u0007", cell.Text);

            cell = row.GetCell(1);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(9, cell._parStart);
            Assert.AreEqual(10, cell._parEnd);
            Assert.AreEqual(93, cell.StartOffset);
            Assert.AreEqual(106, cell.EndOffset);
            Assert.AreEqual("Row 3/Cell 2\u0007", cell.Text);

            cell = row.GetCell(2);
            Assert.AreEqual(1, cell.NumParagraphs);
            Assert.AreEqual(10, cell._parStart);
            Assert.AreEqual(11, cell._parEnd);
            Assert.AreEqual(106, cell.StartOffset);
            Assert.AreEqual(119, cell.EndOffset);
            Assert.AreEqual("Row 3/Cell 3\u0007", cell.Text);
        }
    }

}