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

namespace NPOI.XWPF.Extractor
{
    using System;



    using NUnit.Framework;
    using NPOI.XWPF;
    using NPOI.XWPF.UserModel;
    using NPOI.XWPF.Extractor;

    [TestFixture]
    public class TestExternalEntities
    {

        /**
         * Get text out of the simple file
         *
         * @
         */
        [Test]
        public void TestFile()
        {
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("ExternalEntityInText.docx");
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

            String text = extractor.Text;

            Assert.IsTrue(text.Length > 0);

            // Check contents, they should not contain the text from POI web site After colon!
            Assert.AreEqual("Here should not be the POI web site: \"\"", text.Trim());

            extractor.Close();
        }

    }

}