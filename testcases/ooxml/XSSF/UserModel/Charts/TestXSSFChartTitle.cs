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

namespace TestCases.XSSF.UserModel.Charts
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS.UserModel;
    using NPOI.SS.UserModel.Charts;
    using NPOI.SS.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;

    /**
     * Test Get/set chart title.
     */
    [TestFixture]
    public class TestXSSFChartTitle
    {
        private IWorkbook CreateWorkbookWithChart()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("linechart");
            int NUM_OF_ROWS = 3;
            int NUM_OF_COLUMNS = 10;

            // Create a row and Put some cells in it. Rows are 0 based.
            IRow row;
            ICell cell;
            for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++)
            {
                row = sheet.CreateRow((short)rowIndex);
                for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++)
                {
                    cell = row.CreateCell((short)colIndex);
                    cell.SetCellValue(colIndex * (rowIndex + 1));
                }
            }

            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 0, 5, 10, 15);

            IChart chart = Drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = (/*setter*/LegendPosition.TopRight);

            ILineChartData<double,double> data = chart.ChartDataFactory.CreateLineChartData<double, double>();

            // Use a category axis for the bottom axis.
            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = (/*setter*/AxisCrosses.AutoZero);

            IChartDataSource<double> xs = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
            IChartDataSource<double> ys1 = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
            IChartDataSource<double> ys2 = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));

            data.AddSeries(xs, ys1);
            data.AddSeries(xs, ys2);

            chart.Plot(data, bottomAxis, leftAxis);

            return wb;
        }

        /**
         * Gets the first chart from the named sheet in the workbook.
         */
        private XSSFChart GetChartFromWorkbook(IWorkbook wb, String sheetName)
        {
            ISheet sheet = wb.GetSheet(sheetName);
            if (sheet is XSSFSheet)
            {
                XSSFSheet xsheet = (XSSFSheet)sheet;
                XSSFDrawing drawing = xsheet.CreateDrawingPatriarch() as XSSFDrawing;
                if (drawing != null)
                {
                    List<XSSFChart> charts = drawing.GetCharts();
                    if (charts != null && charts.Count > 0)
                    {
                        return charts[0];
                    }
                }
            }
            return null;
        }

        [Test]
        public void TestNewChart()
        {
            IWorkbook wb = CreateWorkbookWithChart();
            XSSFChart chart = GetChartFromWorkbook(wb, "linechart");
            Assert.IsNotNull(chart);
            Assert.IsNull(chart.Title);
            String myTitle = "My chart title";
            chart.SetTitle(myTitle);
            XSSFRichTextString queryTitle = chart.Title;
            Assert.IsNotNull(queryTitle);
            Assert.AreEqual(myTitle, queryTitle.ToString());
            wb.Close();
        }

        [Test]
        public void TestExistingChartWithTitle()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chartTitle_withTitle.xlsx");
            XSSFChart chart = GetChartFromWorkbook(wb, "Sheet1");
            Assert.IsNotNull(chart);
            XSSFRichTextString originalTitle = chart.Title;
            Assert.IsNotNull(originalTitle);
            String myTitle = "My chart title";
            Assert.IsFalse(myTitle.Equals(originalTitle.ToString()));
            chart.SetTitle(myTitle);
            XSSFRichTextString queryTitle = chart.Title;
            Assert.IsNotNull(queryTitle);
            Assert.AreEqual(myTitle, queryTitle.ToString());
            wb.Close();
        }

        [Test]
        public void TestExistingChartNoTitle()
        {
            IWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chartTitle_noTitle.xlsx");
            XSSFChart chart = GetChartFromWorkbook(wb, "Sheet1");
            Assert.IsNotNull(chart);
            Assert.IsNull(chart.Title);
            String myTitle = "My chart title";
            chart.SetTitle(myTitle);
            XSSFRichTextString queryTitle = chart.Title;
            Assert.IsNotNull(queryTitle);
            Assert.AreEqual(myTitle, queryTitle.ToString());
            wb.Close();
        }
    }
}
