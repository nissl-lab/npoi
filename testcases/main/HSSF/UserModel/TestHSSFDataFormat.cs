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
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;
    using TestCases.SS.UserModel;
using NPOI.SS.UserModel;
    using NPOI.HSSF.UserModel;

    /**
     * Test <c>HSSFPicture</c>.
     *
     * @author Yegor Kozlov (yegor at apache.org)
     */
    [TestClass]
    public class TestHSSFDataFormat:BaseTestDataFormat
    {
        public TestHSSFDataFormat()
            : base(HSSFITestDataProvider.Instance)
        {
            
        }
                    /**
     * [Bug 49928] formatCellValue returns incorrect value for \u00a3 formatted cells
     */
    [TestMethod]
    public void Test49928()
    {
        HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("49928.xls");
        string poundFmt = "\"\u00a3\"#,##0;[Red]\\-\"\u00a3\"#,##0";
        DataFormatter df = new DataFormatter();

        ISheet sheet = wb.GetSheetAt(0);
        ICell cell = sheet.GetRow(0).GetCell(0);
        ICellStyle style = cell.CellStyle;

        // not expected normally, id of a custom format should be greater 
        // than BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX
        short  poundFmtIdx = 6;

        Assert.AreEqual(poundFmt, style.GetDataFormatString());
        Assert.AreEqual(poundFmtIdx, style.DataFormat);
        Assert.AreEqual("\u00a31", df.FormatCellValue(cell));


        IDataFormat dataFormat = wb.CreateDataFormat();
        Assert.AreEqual(poundFmtIdx, dataFormat.GetFormat(poundFmt));
        Assert.AreEqual(poundFmt, dataFormat.GetFormat(poundFmtIdx));
    }
 
    }
}