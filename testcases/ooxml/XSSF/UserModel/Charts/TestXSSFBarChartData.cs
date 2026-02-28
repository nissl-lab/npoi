using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.Linq;

namespace TestCases.XSSF.UserModel.Charts
{
    public class TestXSSFBarChartData
    {
        private static readonly object[][] plotData = new object[][]
        {
            ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"], 
            [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        };
        
        [TestCase(BarGrouping.Stacked, ST_BarGrouping.stacked)]
        [TestCase(BarGrouping.Clustered, ST_BarGrouping.clustered)]
        [TestCase(BarGrouping.Standard, ST_BarGrouping.standard)]
        [TestCase(BarGrouping.PercentStacked, ST_BarGrouping.percentStacked)]
        public void TestSettingBarGrouping(BarGrouping barGrouping, ST_BarGrouping expectedBarGrouping)
        {
            using IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, plotData).Build();
            IDrawing<IShape> drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = drawing.CreateChart(anchor);

            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            IChartAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);

            IBarChartData<string, double> barChartData = chart.ChartDataFactory.CreateBarChartData<string, double>();

            IChartDataSource<string> xs = DataSources.FromStringCellRange(sheet, CellRangeAddress.ValueOf("A1:J1"));
            IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("A2:J2"));
            barChartData.AddSeries(xs, ys);
            
            barChartData.SetBarGrouping(barGrouping);
            
            chart.Plot(barChartData, bottomAxis, leftAxis);
            
            ClassicAssert.IsInstanceOf<XSSFChart>(chart);
            XSSFChart xssfChart = (XSSFChart)chart;
            CT_BarChart ctBarChart = xssfChart.GetCTChart().plotArea.barChart.FirstOrDefault();
            ClassicAssert.NotNull(ctBarChart);
            ClassicAssert.AreEqual(expectedBarGrouping, ctBarChart!.grouping.val);
        }
        
        [Test]
        public void TestBarGroupingBeClusteredWhenNoBarGroupingIsSet()
        {
            using IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, plotData).Build();
            IDrawing<IShape> drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = drawing.CreateChart(anchor);

            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            IChartAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);

            IBarChartData<string, double> barChartData = chart.ChartDataFactory.CreateBarChartData<string, double>();

            IChartDataSource<string> xs = DataSources.FromStringCellRange(sheet, CellRangeAddress.ValueOf("A1:J1"));
            IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("A2:J2"));
            barChartData.AddSeries(xs, ys);
            
            chart.Plot(barChartData, bottomAxis, leftAxis);
            
            ClassicAssert.IsInstanceOf<XSSFChart>(chart);
            XSSFChart xssfChart = (XSSFChart)chart;
            CT_BarChart ctBarChart = xssfChart.GetCTChart().plotArea.barChart.FirstOrDefault();
            ClassicAssert.NotNull(ctBarChart);
            ClassicAssert.AreEqual(ST_BarGrouping.clustered, ctBarChart!.grouping.val);
        }
    }
}