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
    using System;
    using System.IO;
    using NPOI.HSSF.UserModel;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using NUnit.Framework;

    [TestFixture]
    public class TestUnicodeWorkbook
    {

        public TestUnicodeWorkbook()
        {

        }

        /** Tests Bug38230
         *  That a Umlat is written  and then read back.
         *  It should have been written as a compressed unicode.
         * 
         * 
         *
         */
        [Test]
        public void TestUmlatReadWrite()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            //Create a unicode sheet name (euro symbol)
            NPOI.SS.UserModel.ISheet s = wb.CreateSheet("Test");

            IRow r = s.CreateRow(0);
            ICell c = r.CreateCell(1);
            c.SetCellValue(new HSSFRichTextString("\u00e4"));

            //Confirm that the sring will be compressed
            Assert.AreEqual(((HSSFRichTextString)c.RichStringCellValue).UnicodeString.OptionFlags, 0);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);

            //Test the sheetname
            s = wb.GetSheet("Test");
            Assert.IsNotNull(s);

            c = r.GetCell(1);
            Assert.AreEqual(c.RichStringCellValue.String, "\u00e4");
        }

    }
}