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

using NPOI.SS.UserModel;
using System.Collections.Generic;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
namespace TestCases.SS.UserModel
{

    /**
     * Tests of implementation of {@link DataFormat}
     *
     */
    public abstract class BaseTestDataFormat
    {

        protected ITestDataProvider _testDataProvider;

        /**
         * @param testDataProvider an object that provides test data in HSSF / XSSF specific way
         */
        protected BaseTestDataFormat(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }
        [TestMethod]
        public void BaseBuiltinFormats()
        {
            Workbook wb = _testDataProvider.CreateWorkbook();

            DataFormat df = wb.CreateDataFormat();

            List<String> formats = HSSFDataFormat.GetBuiltinFormats();
            int idx = 0;
            foreach (string fmt in formats)
            {
                Assert.AreEqual(idx, df.GetFormat(fmt));
                idx++;
            }

            //default format for new cells is General
            Sheet sheet = wb.CreateSheet();
            Cell cell = sheet.CreateRow(0).CreateCell(0);
            Assert.AreEqual(0, cell.CellStyle.DataFormat);
            Assert.AreEqual("General", cell.CellStyle.GetDataFormatString());

            //create a custom data format
            String customFmt = "#0.00 AM/PM";
            //check it is not in built-in formats
            Assert.AreEqual(-1, HSSFDataFormat.GetBuiltinFormat(customFmt));
            int customIdx = df.GetFormat(customFmt);
            //The first user-defined format starts at 164.
            Assert.IsTrue(customIdx >= HSSFDataFormat.FIRST_USER_DEFINED_FORMAT_INDEX);
            //read and verify the string representation
            Assert.AreEqual(customFmt, df.GetFormat((short)customIdx));
        }


    }
}




