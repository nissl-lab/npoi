namespace TestCases.XSSF.UserModel.Charts
{
    using System;

    using NUnit.Framework;
    using NPOI.SS.UserModel;
    using NPOI.SS.UserModel.Charts;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;

    /**
     * Tests for XSSF Line Charts
     */
    [TestFixture]
    public class TestXSSFLineChartData
    {
        private static readonly object[][] plotData = new object[2][]
        {
            new string[] {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"},
            new object[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 },
        };

        [Test]
        public void TestOneSeriePlot()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, plotData).Build();
            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = Drawing.CreateChart(anchor);

            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            IChartAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);

            ILineChartData<string, double> lineChartData =
                    chart.ChartDataFactory.CreateLineChartData<string, double>();

            IChartDataSource<String> xs = DataSources.FromStringCellRange(sheet, CellRangeAddress.ValueOf("A1:J1"));
            IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("A2:J2"));
            ILineChartSeries<string, double> series = lineChartData.AddSeries(xs, ys);

            Assert.IsNotNull(series);
            Assert.AreEqual(1, lineChartData.GetSeries().Count);
            Assert.IsTrue(lineChartData.GetSeries().Contains(series));

            chart.Plot(lineChartData, bottomAxis, leftAxis);
            wb.Close();
        }
    }
}
