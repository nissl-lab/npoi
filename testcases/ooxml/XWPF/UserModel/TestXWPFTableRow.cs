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

namespace TestCases.XWPF.UserModel
{
    using NPOI.XWPF.UserModel;
    using NUnit.Framework;

    [TestFixture]
    public class TestXWPFTableRow
    {
        [Test]
        public void TestCreateRow()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFTable table = doc.CreateTable(1, 1);
            XWPFTableRow tr = table.CreateRow();
            Assert.IsNotNull(tr);
            doc.Close();
        }

        [Test]
        public void TestSetGetCantSplitRow()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            XWPFTable table = doc.CreateTable(1, 1);
            // table has a single row by default; grab it
            XWPFTableRow tr = table.GetRow(0);
            Assert.IsNotNull(tr);

            // Assert the repeat header is false by default
            bool isCantSplit = tr.IsCantSplitRow;
            Assert.IsFalse(isCantSplit);

            // Repeat the header
            tr.SetCantSplitRow(true);
            isCantSplit = tr.IsCantSplitRow;
            Assert.IsTrue(isCantSplit);

            // Make the header no longer repeating
            tr.SetCantSplitRow(false);
            isCantSplit = tr.IsCantSplitRow;
            Assert.IsFalse(isCantSplit);

            doc.Close();
        }

        [Test]
        public void TestSetGetRepeatHeader()
        {
            // create a table
            XWPFDocument doc = new XWPFDocument();
            XWPFTable table = doc.CreateTable(3, 1);
            // table has a single row by default; grab it
            XWPFTableRow tr = table.GetRow(0);
            Assert.IsNotNull(tr);

            // Assert the repeat header is false by default
            bool isRpt = tr.IsRepeatHeader;
            Assert.IsFalse(isRpt);

            // Repeat the header
            tr.SetRepeatHeader(true);
            isRpt = tr.IsRepeatHeader;
            Assert.IsTrue(isRpt);

            // Make the header no longer repeating
            tr.SetRepeatHeader(false);
            isRpt = tr.IsRepeatHeader;
            Assert.IsFalse(isRpt);

            // If the third row is set to repeat, but not the second,
            // isRepeatHeader should report false because Word will
            // ignore it.
            tr = table.GetRow(2);
            tr.SetRepeatHeader(true);
            isRpt = tr.IsRepeatHeader;
            Assert.IsFalse(isRpt);

            doc.Close();
        }

        // Test that validates the table header value can be parsed from a document
        // generated in Word
        [Test]
        public void TestIsRepeatHeader()
        {
            XWPFDocument doc = XWPFTestDataSamples
                    .OpenSampleDocument("Bug60337.docx");
            XWPFTable table = doc.Tables[0];
            XWPFTableRow tr = table.GetRow(0);
            bool isRpt = tr.IsRepeatHeader;
            Assert.IsTrue(isRpt);

            tr = table.GetRow(1);
            isRpt = tr.IsRepeatHeader;
            Assert.IsFalse(isRpt);

            tr = table.GetRow(2);
            isRpt = tr.IsRepeatHeader;
            Assert.IsFalse(isRpt);
        }


        // Test that validates the table header value can be parsed from a document
        // generated in Word
        [Test]
        public void TestIsCantSplit()
        {
            XWPFDocument doc = XWPFTestDataSamples
                    .OpenSampleDocument("Bug60337.docx");
            XWPFTable table = doc.Tables[0];
            XWPFTableRow tr = table.GetRow(0);
            bool isCantSplit = tr.IsCantSplitRow;
            Assert.IsFalse(isCantSplit);

            tr = table.GetRow(1);
            isCantSplit = tr.IsCantSplitRow;
            Assert.IsFalse(isCantSplit);

            tr = table.GetRow(2);
            isCantSplit = tr.IsCantSplitRow;
            Assert.IsTrue(isCantSplit);
        }
    }
}
