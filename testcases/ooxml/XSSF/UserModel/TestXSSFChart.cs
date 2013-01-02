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

namespace NPOI.XSSF.UserModel
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
            Assert.AreEqual(null, chart.GetTitle());

            chart = (s2.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[(1)];
            Assert.AreEqual("Pie Chart Title Thingy", chart.GetTitle().String);

            chart = (s3.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[(0)];
            Assert.AreEqual("Sheet 3 Chart with Title", chart.GetTitle().String);
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
            Assert.AreEqual(2, d1.GetCharts().Count);
        }
    }

}