/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using System.IO;
using TestCases;
using NPOI.HWPF.UserModel;
using NUnit.Framework;

namespace NPOI.HWPF.SPRM
{
    [TestFixture]
    public class TestSprms
    {
        private static HWPFDocument Reload(HWPFDocument hwpfDocument)
        {
            MemoryStream baos = new MemoryStream();
            hwpfDocument.Write(baos);
            return new HWPFDocument(new MemoryStream(baos.ToArray()));
        }

        /**
         * Test correct Processing of "sprmPItap" (0x6649) and "sprmPFInTable"
         * (0x2416)
         */
        [Test]
        public void TestInnerTable()
        {
            Stream resourceAsStream = POIDataSamples.GetDocumentInstance()
                    .OpenResourceAsStream("innertable.doc");
            HWPFDocument hwpfDocument = new HWPFDocument(resourceAsStream);
            resourceAsStream.Close();

            TestInnerTable(hwpfDocument);
            hwpfDocument = Reload(hwpfDocument);
            TestInnerTable(hwpfDocument);
        }

        private void TestInnerTable(HWPFDocument hwpfDocument)
        {
            Range range = hwpfDocument.GetRange();
            for (int p = 0; p < range.NumParagraphs; p++)
            {
                Paragraph paragraph = range.GetParagraph(p);
                char first = paragraph.Text.ToLower()[0];
                if ('1' <= first && first < '4')
                {
                    Assert.IsTrue(paragraph.IsInTable());
                    Assert.AreEqual(2, paragraph.GetTableLevel());
                }

                if ('a' <= first && first < 'z')
                {
                    Assert.IsTrue(paragraph.IsInTable());
                    Assert.AreEqual(1, paragraph.GetTableLevel());
                }
            }
        }

        /**
         * Test correct Processing of "sprmPJc" by uncompressor
         */
        [Test]
        public void TestSprmPJc()
        {
            Stream resourceAsStream = POIDataSamples.GetDocumentInstance()
                    .OpenResourceAsStream("Bug49820.doc");
            HWPFDocument hwpfDocument = new HWPFDocument(resourceAsStream);
            resourceAsStream.Close();

            Assert.AreEqual(1, hwpfDocument.GetStyleSheet().GetParagraphStyle(8)
                    .GetJustification());

            hwpfDocument = Reload(hwpfDocument);

            Assert.AreEqual(1, hwpfDocument.GetStyleSheet().GetParagraphStyle(8)
                    .GetJustification());

        }
    }
}

