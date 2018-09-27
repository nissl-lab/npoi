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

using System;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
namespace NPOI.XSSF.UserModel.Charts
{
    [TestFixture]
    public class TestXSSFNumberCache
    {
        private static Object[][] plotData = {
	        new object[]{0,      1,       2,       3,       4},
	        new object[]{0, "=B1*2", "=C1*2", "=D1*2", "=E1*4"}
            };
        [Test]
        public void TestFormulaCache()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, plotData).Build();
            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = Drawing.CreateChart(anchor);

            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Bottom);
            IChartAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);

            IScatterChartData scatterChartData =
                chart.ChartDataFactory.CreateScatterChartData();

            DataMarker xMarker = new DataMarker(sheet, CellRangeAddress.ValueOf("A1:E1"));
            DataMarker yMarker = new DataMarker(sheet, CellRangeAddress.ValueOf("A2:E2"));
            IScatterChartSerie serie = scatterChartData.AddSerie(xMarker, yMarker);

            chart.Plot(scatterChartData, bottomAxis, leftAxis);

            XSSFScatterChartData.Serie xssfScatterSerie =
                (XSSFScatterChartData.Serie)serie;
            XSSFNumberCache yCache = xssfScatterSerie.LastCalculatedYCache;

            Assert.AreEqual(5, yCache.PointCount);
            Assert.AreEqual(4.0, yCache.GetValueAt(3), 0.00001);
            Assert.AreEqual(16.0, yCache.GetValueAt(5), 0.00001);
        }


    }
}
