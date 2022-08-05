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
namespace TestCases.SS.Format
{
    using NPOI.SS.Format;
    using NUnit.Framework;
    using SixLabors.ImageSharp.PixelFormats;
    using System;

    [TestFixture]
    public class TestCellFormatResult
    {

        [Test]
        public void TestNullTextRaisesException()
        {
            try
            {
                bool applies = true;
                String text = null;
                Rgb24 textColor = new Rgb24(0, 0, 0);
                CellFormatResult result = new CellFormatResult(applies, text, textColor);
                Assert.Fail("Cannot Initialize CellFormatResult with null text parameter");
            }
            catch (ArgumentException )
            {
                //Expected
            }
        }
    }
}
