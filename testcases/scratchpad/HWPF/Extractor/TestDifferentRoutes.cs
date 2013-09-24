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

using System;
using TestCases.HWPF;
using NPOI.HWPF.UserModel;

using System.Text;
using NPOI.HWPF.Model;
using NPOI.HWPF;
using System.Collections.Generic;
using NUnit.Framework;
namespace TestCases.HWPF.Extractor
{

    /**
     * Test the different routes to extracting text
     *
     * @author Nick Burch (nick at torchbox dot com)
     */
    [TestFixture]
    public class TestDifferentRoutes
    {
        private String[] p_text = new String[] {
			"This is a simple word document\r",
			"\r",
			"It has a number of paragraphs in it\r",
			"\r",
			"Some of them even feature bold, italic and underlined text\r",
			"\r",
			"\r",
			"This bit is in a different font and size\r",
			"\r",
			"\r",
			"This bit features some red text.\r",
			"\r",
			"\r",
			"It is otherwise very very boring.\r"
	};

        private HWPFDocument doc;
        [SetUp]
        public void SetUp()
        {
            doc = HWPFTestDataSamples.OpenSampleFile("test2.doc");
        }

        /**
         * Test model based extraction
         */
        [Test]
        public void TestExtractFromModel()
        {
            Range r = doc.GetRange();

            String[] text = new String[r.NumParagraphs];
            for (int i = 0; i < r.NumParagraphs; i++)
            {
                Paragraph p = r.GetParagraph(i);
                text[i] = p.Text;
            }

            Assert.AreEqual(p_text.Length, text.Length);
            for (int i = 0; i < p_text.Length; i++)
            {
                Assert.AreEqual(p_text[i], text[i]);
            }
        }

        /**
         * Test textPieces based extraction
         */
        [Test]
        public void TestExtractFromTextPieces()
        {
            StringBuilder textBuf = new StringBuilder();

            List<TextPiece> textPieces = doc.TextTable.TextPieces;
            foreach(TextPiece piece in textPieces)
            {
                String encoding = "Windows-1252";
                if (piece.IsUnicode)
                {
                    encoding = "UTF-16LE";
                }
                String text = Encoding.GetEncoding(encoding).GetString(piece.RawBytes);
                textBuf.Append(text);
            }

            StringBuilder exp = new StringBuilder();
            for (int i = 0; i < p_text.Length; i++)
            {
                exp.Append(p_text[i]);
            }
            Assert.AreEqual(exp.ToString(), textBuf.ToString());
        }
    }
}
