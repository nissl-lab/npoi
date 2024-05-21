using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
/*
*  ====================================================================
*    Licensed to the Apache Software Foundation (ASF) under one or more
*    contributor license agreements.  See the NOTICE file distributed with
*    this work for additional information regarding copyright ownership.
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
using TestCases.SS.UserModel;
namespace TestCases.XSSF.UserModel
{
    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestXSSFConditionalFormatting : BaseTestConditionalFormatting
    {
        protected override void AssertColour(String hexExpected, IColor actual)
        {
            Assert.IsNotNull(actual, "Colour must be given");
            XSSFColor colour = (XSSFColor)actual;
            if (hexExpected.Length == 8)
            {
                Assert.AreEqual(hexExpected, colour.ARGBHex);
            }
            else
            {
                Assert.AreEqual(hexExpected, colour.ARGBHex.Substring(2));
            }
        }

        public TestXSSFConditionalFormatting()
            : base(XSSFITestDataProvider.instance)
        {

        }
        [Test]
        public void TestRead()
        {
            this.TestRead("WithConditionalFormatting.xlsx");
        }
        [Test]
        public void TestReadOffice2007()
        {
            // TODO Bring the XSSF support up to the same level
            TestReadOffice2007("NewStyleConditionalFormattings.xlsx");
        }
    }
}