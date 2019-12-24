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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using NPOI.HSSF.Model;
    using System.Collections.Generic;

    /**
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestReadWriteChart
    {

        /**
         * In the presence of a chart we need to make sure BOF/EOF records still exist.
         */
        [Test]
        public void TestBOFandEOFRecords()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("SimpleChart.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            IRow firstRow = sheet.GetRow(0);
            ICell firstCell = firstRow.GetCell(0);

            //System.out.println("first assertion for date");
            Assert.AreEqual(new DateTime(2000, 1, 1, 10, 51, 2),
                         DateUtil.GetJavaDate(firstCell.NumericCellValue, false));
            IRow row = sheet.CreateRow(15);
            ICell cell = row.CreateCell(1);

            cell.SetCellValue(22);
            InternalSheet newSheet = ((HSSFSheet)workbook.GetSheetAt(0)).Sheet;
            IList<RecordBase> records = newSheet.Records;

            //System.out.println("BOF Assertion");
            Assert.IsTrue(records[0] is BOFRecord);
            //System.out.println("EOF Assertion");
            Assert.IsTrue(records[records.Count - 1] is EOFRecord);

            workbook.Close();
        }
    }
}
