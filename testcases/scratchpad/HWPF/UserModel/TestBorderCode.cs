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

using NPOI.HWPF.UserModel;
using NPOI.HWPF;

using System;
using NUnit.Framework;
namespace TestCases.HWPF.UserModel
{

    /**
     * API for BorderCode, see Bugzill 49919
     */
    [TestFixture]
    public class TestBorderCode
    {

        private int pos = 0;
        private Range range;

        [Test]
        public void Test()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug49919.doc");
            range = doc.GetRange();

            Paragraph par = FindParagraph("Paragraph normal\r");
            Assert.AreEqual(0, par.GetLeftBorder().BorderType);
            Assert.AreEqual(0, par.GetRightBorder().BorderType);
            Assert.AreEqual(0, par.GetTopBorder().BorderType);
            Assert.AreEqual(0, par.GetBottomBorder().BorderType);

            par = FindParagraph("Paragraph with border\r");
            Assert.AreEqual(18, par.GetLeftBorder().BorderType);
            Assert.AreEqual(17, par.GetRightBorder().BorderType);
            Assert.AreEqual(18, par.GetTopBorder().BorderType);
            Assert.AreEqual(17, par.GetBottomBorder().BorderType);
            Assert.AreEqual(15, par.GetLeftBorder().Color);

            par = FindParagraph("Paragraph with red border\r");
            Assert.AreEqual(1, par.GetLeftBorder().BorderType);
            Assert.AreEqual(1, par.GetRightBorder().BorderType);
            Assert.AreEqual(1, par.GetTopBorder().BorderType);
            Assert.AreEqual(1, par.GetBottomBorder().BorderType);
            Assert.AreEqual(6, par.GetLeftBorder().Color);

            par = FindParagraph("Paragraph with bordered words.\r");
            Assert.AreEqual(0, par.GetLeftBorder().BorderType);
            Assert.AreEqual(0, par.GetRightBorder().BorderType);
            Assert.AreEqual(0, par.GetTopBorder().BorderType);
            Assert.AreEqual(0, par.GetBottomBorder().BorderType);

            Assert.AreEqual(3, par.NumCharacterRuns);
            CharacterRun chr = par.GetCharacterRun(0);
            Assert.AreEqual(0, chr.GetBorder().BorderType);
            chr = par.GetCharacterRun(1);
            Assert.AreEqual(1, chr.GetBorder().BorderType);
            Assert.AreEqual(0, chr.GetBorder().Color);
            chr = par.GetCharacterRun(2);
            Assert.AreEqual(0, chr.GetBorder().BorderType);

            while (pos < range.NumParagraphs - 1)
            {
                par = range.GetParagraph(pos++);
                if (par.IsInTable())
                    break;
            }

            Assert.AreEqual(true, par.IsInTable());
            Table tab = range.GetTable(par);

            // Border could be defined for the entire row, or for each cell, with the same visual appearance.
            Assert.AreEqual(2, tab.NumRows);
            TableRow row = tab.GetRow(0);
            Assert.AreEqual(1, row.GetLeftBorder().BorderType);
            Assert.AreEqual(1, row.GetRightBorder().BorderType);
            Assert.AreEqual(1, row.GetTopBorder().BorderType);
            Assert.AreEqual(1, row.GetBottomBorder().BorderType);

            Assert.AreEqual(2, row.NumCells());
            TableCell cell = row.GetCell(1);
            Assert.AreEqual(3, cell.GetBrcTop().BorderType);

            row = tab.GetRow(1);
            cell = row.GetCell(0);
            // 255 Clears border inherited from row
            Assert.AreEqual(255, cell.GetBrcBottom().BorderType);
        }

        private Paragraph FindParagraph(String expectedText)
        {
            while (pos < range.NumParagraphs - 1)
            {
                Paragraph par = range.GetParagraph(pos);
                pos++;
                if (par.Text.Equals(expectedText))
                    return par;
            }

            Assert.Fail("Expected paragraph not found");

            // should never come here
            throw null;
        }

    }
}

