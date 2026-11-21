namespace TestCases.XSSF.UserModel.Charts
{
    using System;

    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;
    using NPOI.XDDF.UserModel.Chart;

    /**
     * Tests for XSSF Line Charts
     */
    [TestFixture]
    public class TestXDDFLineChartData
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
            var Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            var chart = Drawing.CreateChart(anchor);

            var bottomAxis = chart.CreateCategoryAxis(AxisPosition.Bottom);
            var leftAxis = chart.CreateValueAxis(AxisPosition.Left);

            var lineChartData =
                    chart.CreateData<string, double>(ChartTypes.LINE, bottomAxis, leftAxis);

            var xs = XDDFDataSourcesFactory.FromStringCellRange(sheet, CellRangeAddress.ValueOf("A1:J1"));
            var ys = XDDFDataSourcesFactory.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("A2:J2"));
            var series = lineChartData.AddSeries(xs, ys);

            ClassicAssert.IsNotNull(series);
            ClassicAssert.AreEqual(1, lineChartData.GetSeries().Count);
            ClassicAssert.IsTrue(lineChartData.GetSeries().Contains(series));

            chart.Plot(lineChartData);
            wb.Close();
        }
    }
}
