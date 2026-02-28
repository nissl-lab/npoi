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
using NUnit.Framework;

namespace TestCases.HWPF.UserModel
{

    /**
     * API for Processing of symbols, see Bugzilla 49908
     */
    [TestFixture]
    public class TestRangeSymbols
    {
        [Test]
        public void Test()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("Bug49908.doc");

            Range range = doc.GetRange();

            Assert.IsTrue(range.NumCharacterRuns >= 2);
            CharacterRun chr = range.GetCharacterRun(0);
            Assert.AreEqual(false, chr.IsSymbol());

            chr = range.GetCharacterRun(1);
            Assert.AreEqual(true, chr.IsSymbol());
            Assert.AreEqual("\u0028", chr.Text);
            Assert.AreEqual("Wingdings", chr.GetSymbolFont().GetMainFontName());
            Assert.AreEqual(0xf028, chr.GetSymbolCharacter());
        }

    }
}