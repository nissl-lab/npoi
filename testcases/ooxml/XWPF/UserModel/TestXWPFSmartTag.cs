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
namespace NPOI.XWPF.UserModel
{
    using NPOI.XWPF;
    using NUnit.Framework;

    /**
     * Tests for Reading SmartTags from Word docx.
     *
     * @author  Fabian Lange
     */
    [TestFixture]
    public class TestXWPFSmartTag
    {
        [Test]
        public void TestSmartTags()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("smarttag-snippet.docx");
            XWPFParagraph p = doc.GetParagraphArray(0);
            //About NPOI: because the serializer bug(the CT_Run contains whitespace will discard the whitespace),
            //Text is "CarnegieMellonUniversitySchool of Computer Science"
            Assert.IsTrue(p.Text.Contains("Carnegie Mellon University School of Computer Science"));
            p = doc.GetParagraphArray(2);
            Assert.IsTrue(p.Text.Contains("Alice's Adventures"));
        }
    }

}