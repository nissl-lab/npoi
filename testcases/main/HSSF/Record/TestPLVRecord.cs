/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using NUnit.Framework;
namespace TestCases.HSSF.Record
{
    /**
     * Verify that presence of PLV record doesn't break data
     * validation, bug #53972:
     * https://issues.apache.org/bugzilla/show_bug.cgi?id=53972
     *
     * @author Andrew Novikov
     */
    [TestFixture]
    public class TestPLVRecord
    {
        private static String DV_DEFINITION = "$A$1:$A$5";
        private static String XLS_FILENAME = "53972.xls";
        private static String SHEET_NAME = "S2";
        [Test]
        public void TestPLVRecord1()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(XLS_FILENAME);
            HSSFWorkbook workbook = new HSSFWorkbook(is1);

            CellRangeAddressList cellRange = new CellRangeAddressList(0, 0, 1, 1);
            IDataValidationConstraint constraint = DVConstraint.CreateFormulaListConstraint(DV_DEFINITION);
            HSSFDataValidation dataValidation = new HSSFDataValidation(cellRange, constraint);

            // This used to throw an error before
            try
            {
                workbook.GetSheet(SHEET_NAME).AddValidationData(dataValidation);
            }
            catch (InvalidOperationException)
            {
                Assert.Fail("Identified bug 53972, PLV record breaks addDataValidation()");
            }
        }
    }
}