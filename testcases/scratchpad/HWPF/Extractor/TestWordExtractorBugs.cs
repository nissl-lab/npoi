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

using NPOI.HWPF.Extractor;
using NUnit.Framework;

namespace TestCases.HWPF.Extractor
{

    /**
     * Tests for bugs with the WordExtractor
     *
     * @author Nick Burch (nick at torchbox dot com)
     */
    [TestFixture]
    public class TestWordExtractorBugs
    {
        [Test]
        public void TestProblemMetadata()
        {
            WordExtractor extractor =
                new WordExtractor(POIDataSamples.GetDocumentInstance().OpenResourceAsStream("ProblemExtracting.doc"));

            // Check it gives text without error
            string text=extractor.Text;
            string[] paratext=extractor.ParagraphText;
            string textfrompieces=extractor.TextFromPieces;
        }
    }

}