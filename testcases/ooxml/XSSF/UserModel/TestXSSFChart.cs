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

using NUnit.Framework;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using NPOI.SS.UserModel.Charts;
using NPOI.XSSF.UserModel;
using NPOI.XSSF;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System.Linq;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFChart
    {
        [Test]
        public void TestGetAccessors()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithThreeCharts.xlsx");
            XSSFSheet s1 = (XSSFSheet)wb.GetSheetAt(0);
            XSSFSheet s2 = (XSSFSheet)wb.GetSheetAt(1);
            XSSFSheet s3 = (XSSFSheet)wb.GetSheetAt(2);

            Assert.AreEqual(0, s1.GetRelations().Count);
            Assert.AreEqual(1, s2.GetRelations().Count);
            Assert.AreEqual(1, s3.GetRelations().Count);

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        [Test]
        public void TestGetCharts()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("WithThreeCharts.xlsx");

            XSSFSheet s1 = (XSSFSheet)wb.GetSheetAt(0);
            XSSFSheet s2 = (XSSFSheet)wb.GetSheetAt(1);
            XSSFSheet s3 = (XSSFSheet)wb.GetSheetAt(2);

            Assert.AreEqual(0, (s1.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);
            Assert.AreEqual(2, (s2.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);
            Assert.AreEqual(1, (s3.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);

            // Check the titles
            XSSFChart chart = (s2.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[(0)];
            Assert.AreEqual(null, chart.Title);

            chart = (s2.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[(1)];
            Assert.AreEqual("Pie Chart Title Thingy", chart.Title.String);

            chart = (s3.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[(0)];
            Assert.AreEqual("Sheet 3 Chart with Title", chart.Title.String);

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }

        [Test]
        public void TestAddChartsToNewWorkbook()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet s1 = (XSSFSheet)wb.CreateSheet();
            XSSFDrawing d1 = (XSSFDrawing)s1.CreateDrawingPatriarch();
            XSSFClientAnchor a1 = new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            XSSFChart c1 = (XSSFChart)d1.CreateChart(a1);

            Assert.AreEqual(1, d1.GetCharts().Count);
            
            Assert.IsNotNull(c1.GetGraphicFrame());
            Assert.IsNotNull(c1.GetOrCreateLegend());

            XSSFClientAnchor a2 = new XSSFClientAnchor(0, 0, 0, 0, 1, 11, 10, 60);
            XSSFChart c2 = (XSSFChart)d1.CreateChart(a2);

            Assert.IsNotNull(c2);
            Assert.AreEqual(2, d1.GetCharts().Count);
            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }

        [Test]
        public void TestRemoveChart()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            XSSFDrawing drawingPatriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFClientAnchor anchor1 = new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            XSSFChart chart = (XSSFChart)drawingPatriarch.CreateChart(anchor1);
            XSSFClientAnchor anchor2 = new XSSFClientAnchor(0, 0, 0, 0, 1, 11, 10, 60);
            _ = (XSSFChart)drawingPatriarch.CreateChart(anchor2);

            Assert.AreEqual(2, drawingPatriarch.GetCharts().Count);

            drawingPatriarch.RemoveChart(chart);

            Assert.AreEqual(1, drawingPatriarch.GetCharts().Count);
        }

        [Test]
        public void TestRemoveOneOfTwoChartsAndAddNewOne()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet();
            XSSFDrawing drawingPatriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFClientAnchor anchor1 = new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            XSSFChart chart = (XSSFChart)drawingPatriarch.CreateChart(anchor1);
            XSSFClientAnchor anchor2 = new XSSFClientAnchor(0, 0, 0, 0, 1, 11, 10, 60);
            _ = (XSSFChart)drawingPatriarch.CreateChart(anchor2);

            Assert.AreEqual(2, drawingPatriarch.GetCharts().Count);

            drawingPatriarch.RemoveChart(chart);

            Assert.AreEqual(1, drawingPatriarch.GetCharts().Count);

            XSSFClientAnchor a3 = new XSSFClientAnchor(0, 0, 0, 0, 1, 111, 10, 90);
            _ = (XSSFChart)drawingPatriarch.CreateChart(a3);

            Assert.AreEqual(2, drawingPatriarch.GetCharts().Count);
        }

        [Test]
        public void TestGetChartAxisBug57362()
        {
            //Load existing excel with some chart on it having primary and secondary axis.
            IWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("57362.xlsx");
            ISheet sh = workbook.GetSheetAt(0);
            XSSFSheet xsh = (XSSFSheet)sh;
            XSSFDrawing Drawing = xsh.CreateDrawingPatriarch() as XSSFDrawing;
            XSSFChart chart = Drawing.GetCharts()[(0)];

            List<IChartAxis> axisList = chart.GetAxis();

            Assert.AreEqual(4, axisList.Count);
            Assert.IsNotNull(axisList[(0)]);
            Assert.IsNotNull(axisList[(1)]);
            Assert.IsNotNull(axisList[(2)]);
            Assert.IsNotNull(axisList[(3)]);
        }

    }

}