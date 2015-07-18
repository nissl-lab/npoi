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

namespace TestCases.SS.UserModel
{
    using System;
    using NUnit.Framework;
    using TestCases.SS;
    using NPOI.SS.UserModel;

    /**
     * Class for combined testing of XML-specific functionality of 
     * {@link XSSFCell} and {@link SXSSFCell}.
     * 
     *  Any test that is applicable for {@link HSSFCell} as well should go into
     *  the common base class {@link BaseTestCell}.
     */
    [TestFixture]
    public abstract class BaseTestXCell : BaseTestCell
    {
        protected BaseTestXCell(ITestDataProvider testDataProvider)
            : base(testDataProvider)
        {
            ;
        }
        [Test]
        public void TestXmlEncoding()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet();
            IRow row = sh.CreateRow(0);
            ICell cell = row.CreateCell(0);
            String sval = "\u0000\u0002\u0012<>\t\n\u00a0 &\"POI\'\u2122";
            cell.SetCellValue(sval);

            wb = _testDataProvider.WriteOutAndReadBack(wb);

            // invalid characters are Replaced with question marks
            Assert.AreEqual("???<>\t\n\u00a0 &\"POI\'\u2122", wb.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
        }
    }

}