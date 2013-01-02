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

using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using NPOI.HWPF.UserModel;
using NPOI.HWPF;
namespace TestCases.HWPF.UserModel
{

    [TestClass]
    public class TestBug46610
    {
        [TestMethod]
        public void TestUtf()
        {
            RunExtract("Bug46610_1.doc");
        }
        [TestMethod]
        public void TestUtf2()
        {
            RunExtract("Bug46610_2.doc");
        }
        [TestMethod]
        public void TestExtraction()
        {
            String text = RunExtract("Bug46610_3.doc");
            Assert.IsTrue(text.Contains("\u0421\u0412\u041e\u042e"));
        }

        private static String RunExtract(String sampleName)
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile(sampleName);
            StringBuilder out1 = new StringBuilder();

            Range globalRange = doc.GetRange();
            for (int i = 0; i < globalRange.NumParagraphs; i++)
            {
                Paragraph p = globalRange.GetParagraph(i);
                out1.Append(p.Text);
                out1.Append("\n");
                for (int j = 0; j < p.NumCharacterRuns; j++)
                {
                    CharacterRun characterRun = p.GetCharacterRun(j);
                    string a = characterRun.Text;
                }
            }
            return out1.ToString();
        }
    }

}