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

using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.XSSF.UserModel;

namespace TestCases.XSSF.UserModel.Charts
{
    /**
     * Tests ChartLegend
     */
    [TestFixture]
    public class TestXSSFChartLegend
    {
        [Test]
        public void TestLegendPositionAccessMethods()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = Drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();

            legend.Position = LegendPosition.TopRight;
            Assert.AreEqual(LegendPosition.TopRight, legend.Position);

            wb.Close();
        }

        [Test]
        public void Test_setOverlay_defaultChartLegend_expectOverlayInitialValueSetToFalse()
        {
            // Arrange
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = Drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();

            // Act

            // Assert
            Assert.IsFalse(legend.IsOverlay);

            wb.Close();
        }

        [Test]
        public void Test_setOverlay_chartLegendSetToTrue_expectOverlayInitialValueSetToTrue()
        {
            // Arrange
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = Drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();

            // Act
            legend.IsOverlay = (/*setter*/true);

            // Assert
            Assert.IsTrue(legend.IsOverlay);

            wb.Close();
        }
    }
}
