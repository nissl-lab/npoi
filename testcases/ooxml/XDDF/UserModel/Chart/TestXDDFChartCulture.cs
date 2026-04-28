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
                row.CreateCell(0).SetCellValue(1200.5); // Numeric category
                row.CreateCell(1).SetCellValue(2400.1); // Numeric value

                XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                XSSFClientAnchor anchor = (XSSFClientAnchor)drawing.CreateAnchor(0, 0, 0, 0, 2, 2, 10, 10);
                XSSFChart chart = (XSSFChart)drawing.CreateChart(anchor);

                XDDFCategoryAxis bottomAxis = chart.CreateCategoryAxis(AxisPosition.Bottom);
                XDDFValueAxis leftAxis = chart.CreateValueAxis(AxisPosition.Left);

                XDDFChartData<double, double> data = chart.CreateData<double, double>(ChartTypes.BAR, bottomAxis, leftAxis);
                
                var cat = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new NPOI.SS.Util.CellRangeAddress(1, 1, 0, 0));
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

                    // Robust check: Parse XML and look for <v> elements in the chart namespace
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(xml);
                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                    // OpenXML Chart namespace
                    nsmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

                    XmlNodeList vElements = doc.SelectNodes("//c:v", nsmgr);
                    bool foundCorrectCat = false;
                    bool foundIncorrectCat = false;
                    bool foundCorrectVal = false;
                    bool foundIncorrectVal = false;

                    foreach (XmlNode v in vElements)
                    {
                        if (v.InnerText == "1200.5") foundCorrectCat = true;
                        if (v.InnerText == "1200,5") foundIncorrectCat = true;
                        if (v.InnerText == "2400.1") foundCorrectVal = true;
                        if (v.InnerText == "2400,1") foundIncorrectVal = true;
                    }

                    ClassicAssert.IsTrue(foundCorrectCat, "Category '1200.5' should be present in XML with a dot separator.");
                    ClassicAssert.IsFalse(foundIncorrectCat, "Category '1200,5' should NOT be present in XML with a comma separator.");
                    ClassicAssert.IsTrue(foundCorrectVal, "Value '2400.1' should be present in XML with a dot separator.");
                    ClassicAssert.IsFalse(foundIncorrectVal, "Value '2400,1' should NOT be present in XML with a comma separator.");

                    // Also check the embedded workbook cells (FillSheet)
                    XSSFWorkbook embeddedWb = chart.GetWorkbook();
                    ISheet embeddedSheet = embeddedWb.GetSheetAt(0);
                    IRow embeddedRow = embeddedSheet.GetRow(1); // Row 1 has data
                    ICell categoryCell = embeddedRow.GetCell(0); // Col 0 is category

                    ClassicAssert.AreEqual(CellType.Numeric, categoryCell.CellType, "Category cell in embedded workbook should be numeric");
                    ClassicAssert.AreEqual(1200.5, categoryCell.NumericCellValue, 0.0001, "Category cell value should match");
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = originalCulture;
            }
        }

        [Test]
        public void TestDateTimeCategoryPreservesCulture()
        {
            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
            try
            {
                // Set culture to one that uses dots for dates and commas for decimals
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                DateTime testDate = new DateTime(2026, 4, 28);
                string expectedDateString = testDate.ToString(); // "28.04.2026 ..." in de-DE

                XSSFWorkbook wb = new XSSFWorkbook();
                XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("Sheet1");
                
                IRow row = sheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("Date");
                row.CreateCell(1).SetCellValue("Value");
                
                row = sheet.CreateRow(1);
                row.CreateCell(0).SetCellValue(testDate);
                row.CreateCell(1).SetCellValue(1200.5);

                XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                XSSFClientAnchor anchor = (XSSFClientAnchor)drawing.CreateAnchor(0, 0, 0, 0, 2, 2, 10, 10);
                XSSFChart chart = (XSSFChart)drawing.CreateChart(anchor);

                XDDFCategoryAxis bottomAxis = chart.CreateCategoryAxis(AxisPosition.Bottom);
                XDDFValueAxis leftAxis = chart.CreateValueAxis(AxisPosition.Left);

                XDDFChartData<string, double> data = chart.CreateData<string, double>(ChartTypes.BAR, bottomAxis, leftAxis);
                
                // Use a manual data source to ensure we pass a string that should be preserved
                var cat = XDDFDataSourcesFactory.FromArray(new string[] { expectedDateString }, null);
                var val = XDDFDataSourcesFactory.FromNumericCellRange(sheet, new NPOI.SS.Util.CellRangeAddress(1, 1, 1, 1));
                
                data.AddSeries(cat, val);
                chart.Plot(data);

                using (MemoryStream ms = new MemoryStream())
                {
                    wb.Write(ms);
                    
                    var chartSpace = chart.GetCTChartSpace();
                    string xml;
                    using (var xmlStream = new MemoryStream())
                    {
                        chartSpace.Save(xmlStream);
                        xml = System.Text.Encoding.UTF8.GetString(xmlStream.ToArray());
                    }

                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(xml);
                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                    nsmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

                    // Check string cache (labels)
                    XmlNodeList vElements = doc.SelectNodes("//c:v", nsmgr);
                    bool foundDate = false;
                    bool foundValue = false;

                    foreach (XmlNode v in vElements)
                    {
                        if (v.InnerText == expectedDateString) foundDate = true;
                        if (v.InnerText == "1200.5") foundValue = true;
                    }

                    ClassicAssert.IsTrue(foundDate, "DateTime label should use culture-sensitive formatting: " + expectedDateString);
                    ClassicAssert.IsTrue(foundValue, "Numeric value should still use dots in XML: 1200.5");
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = originalCulture;
            }
        }
    }
}
