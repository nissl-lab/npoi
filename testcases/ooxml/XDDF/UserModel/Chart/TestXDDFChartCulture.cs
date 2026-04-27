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

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XDDF.UserModel.Chart;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.Xml;

namespace TestCases.XDDF.UserModel.Chart
{
    [TestFixture]
    public class TestXDDFChartCulture
    {
        [Test]
        public void TestNumericConversionInNonUSCulture()
        {
            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
            try
            {
                // Set culture to one that uses comma as decimal separator
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

                XSSFWorkbook wb = new XSSFWorkbook();
                XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("Sheet1");
                
                // Create some data
                IRow row = sheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("Category");
                row.CreateCell(1).SetCellValue("Value");
                
                row = sheet.CreateRow(1);
                row.CreateCell(0).SetCellValue("A");
                row.CreateCell(1).SetCellValue(1200.5); // 1200.5 should be 1200,5 in de-DE

                XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                XSSFClientAnchor anchor = (XSSFClientAnchor)drawing.CreateAnchor(0, 0, 0, 0, 2, 2, 10, 10);
                XSSFChart chart = (XSSFChart)drawing.CreateChart(anchor);

                XDDFCategoryAxis bottomAxis = chart.CreateCategoryAxis(AxisPosition.Bottom);
                XDDFValueAxis leftAxis = chart.CreateValueAxis(AxisPosition.Left);

                XDDFChartData<string, double> data = chart.CreateData<string, double>(ChartTypes.BAR, bottomAxis, leftAxis);
                
                var cat = XDDFDataSourcesFactory.FromStringCellRange(sheet, new NPOI.SS.Util.CellRangeAddress(1, 1, 0, 0));
                var val = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new NPOI.SS.Util.CellRangeAddress(1, 1, 1, 1));
                
                data.AddSeries(cat, val);
                chart.Plot(data);

                // Export to XML and check values
                using (MemoryStream ms = new MemoryStream())
                {
                    wb.Write(ms);
                    
                    // We need to look into the chart XML. 
                    // Instead of full unzip, we can inspect the XDDFChart's internal CT_ChartSpace
                    var chartSpace = chart.GetCTChartSpace();
                    string xml;
                    using (var xmlStream = new MemoryStream())
                    {
                        chartSpace.Save(xmlStream);
                        xml = System.Text.Encoding.UTF8.GetString(xmlStream.ToArray());
                    }

                    // Check if the value 1200.5 is present with a dot, not a comma
                    // In XML it should look like <c:v>1200.5</c:v>
                    ClassicAssert.IsTrue(xml.Contains("<c:v>1200.5</c:v>"), "XML should contain '1200.5' with a dot separator, found: " + xml);
                    ClassicAssert.IsFalse(xml.Contains("<c:v>1200,5</c:v>"), "XML should NOT contain '1200,5' with a comma separator");
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = originalCulture;
            }
        }
    }
}
