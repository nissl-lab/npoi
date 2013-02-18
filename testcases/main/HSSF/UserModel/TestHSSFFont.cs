/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace TestCases.HSSF.UserModel
{
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;

    using NPOI.SS.UserModel;
    using TestCases.SS.UserModel;

    /**
     * Test <c>HSSFPicture</c>.
     *
     * @author Yegor Kozlov (yegor at apache.org)
     */
    [TestFixture]
    public class TestHSSFFont:BaseTestFont
    {
        public TestHSSFFont()
            : base(HSSFITestDataProvider.Instance)
        {
            
        }

        [Test]
        public void TestDefaultFont()
        {
            BaseTestDefaultFont(HSSFFont.FONT_ARIAL, (short)200, (short)FontColor.Normal);
        }
    }
}