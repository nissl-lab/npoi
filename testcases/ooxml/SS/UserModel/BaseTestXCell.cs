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
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ISheet sh = wb1.CreateSheet();
            IRow row = sh.CreateRow(0);
            ICell cell = row.CreateCell(0);
            String sval = "\u0000\u0002\u0012<>\t\n\u00a0 &\"POI\'\u2122";
            cell.SetCellValue(sval);

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);
            wb1.Close();

            // invalid characters are Replaced with question marks
            Assert.AreEqual("???<>\t\n\u00a0 &\"POI\'\u2122", wb2.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
            wb2.Close();
        }

        [Test]
        public void TestSetNullValues()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICell cell = wb.CreateSheet("test").CreateRow(0).CreateCell(0);

            //cell.SetCellValue((DateTime?)null);
            //cell.setCellValue((Date)null);
            cell.SetCellValue((String)null);
            cell.SetCellValue((IRichTextString)null);
            cell.SetCellValue((String)null);

            wb.Close();
        }
    }

}