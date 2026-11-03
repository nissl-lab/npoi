using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XDDF.UserModel.Chart;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace TestCases.XSSF.UserModel.Charts
{
    public class TestXDDFChartAxisOrder
    {
        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        private static readonly object[][] plotData = new object[][]
        {
            ["A", "B", "C", "D", "E"],
            [1, 2, 3, 4, 5],
            [50, 40, 30, 20, 10],
        };

        // Excel assigns chart groups to the primary and secondary axis pairs by the document order
        // of the axis elements. Its own order for a combination chart is primary catAx, primary valAx,
        // secondary valAx, secondary catAx; writing the axes grouped by type (valAx, valAx, catAx,
        // catAx) puts both groups on the secondary pair.
        [Test]
        public void TestDualAxisChartKeepsAxisInsertionOrderAcrossRoundTrips()
        {
            using XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = new SheetBuilder(wb, plotData).Build();
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFChart chart = (XSSFChart)drawing.CreateChart(drawing.CreateAnchor(0, 0, 0, 0, 1, 4, 10, 30));

            var xs = XDDFDataSourcesFactory.FromStringCellRange(sheet, CellRangeAddress.ValueOf("A1:E1"));
            var primaryValues = XDDFDataSourcesFactory.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("A2:E2"));
            var secondaryValues = XDDFDataSourcesFactory.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("A3:E3"));

            XDDFCategoryAxis primaryCategory = chart.CreateCategoryAxis(AxisPosition.Bottom);
            XDDFValueAxis primaryValue = chart.CreateValueAxis(AxisPosition.Left);
            var bar = chart.CreateData<string, double>(ChartTypes.BAR, primaryCategory, primaryValue);
            bar.AddSeries(xs, primaryValues).SetTitle("Primary");
            chart.Plot(bar);

            XDDFValueAxis secondaryValue = chart.CreateValueAxis(AxisPosition.Right);
            XDDFCategoryAxis secondaryCategory = chart.CreateCategoryAxis(AxisPosition.Bottom);
            secondaryCategory.IsVisible = false;
            secondaryValue.Crosses = AxisCrosses.Max;
            secondaryCategory.CrossAxis(secondaryValue);
            secondaryValue.CrossAxis(secondaryCategory);
            var line = chart.CreateData<string, double>(ChartTypes.LINE, secondaryCategory, secondaryValue);
            line.AddSeries(xs, secondaryValues).SetTitle("Secondary");
            chart.Plot(line);

            string[] expected = ["catAx", "valAx", "valAx", "catAx"];
            CollectionAssert.AreEqual(expected, AxisTypes(chart.GetCTChart().plotArea));

            using XSSFWorkbook readBack = XSSFTestDataSamples.WriteOutAndReadBack(wb);
            XSSFChart readChart = SingleChart(readBack);
            CollectionAssert.AreEqual(expected, AxisElementNames(readChart));
            CollectionAssert.AreEqual(expected, AxisTypes(readChart.GetCTChart().plotArea));

            // The parsed order must survive a second write, not only the one that created it.
            using XSSFWorkbook readTwice = XSSFTestDataSamples.WriteOutAndReadBack(readBack);
            CollectionAssert.AreEqual(expected, AxisElementNames(SingleChart(readTwice)));
        }

        [Test]
        public void TestAxesAddedDirectlyToTypedListsAreStillWritten()
        {
            CT_PlotArea plotArea = new CT_PlotArea();
            plotArea.AddNewCatAx();
            plotArea.valAx = new List<CT_ValAx> { new CT_ValAx() };
            plotArea.AddNewDateAx();

            CollectionAssert.AreEqual(new[] { "catAx", "dateAx", "valAx" }, AxisTypes(plotArea));

            plotArea.catAx.Clear();
            CollectionAssert.AreEqual(new[] { "dateAx", "valAx" }, AxisTypes(plotArea));
        }

        private static List<string> AxisTypes(CT_PlotArea plotArea)
        {
            return plotArea.GetAxesInDocumentOrder().Select(axis => axis switch
            {
                CT_CatAx => "catAx",
                CT_ValAx => "valAx",
                CT_DateAx => "dateAx",
                CT_SerAx => "serAx",
                _ => axis.GetType().Name,
            }).ToList();
        }

        private static List<string> AxisElementNames(XSSFChart chart)
        {
            XmlDocument xml = ChartXml(chart, out XmlNamespaceManager ns);
            return xml.SelectSingleNode("//c:plotArea", ns).ChildNodes
                .OfType<XmlElement>()
                .Select(element => element.LocalName)
                .Where(name => name is "catAx" or "valAx" or "dateAx" or "serAx")
                .ToList();
        }

        private static XmlDocument ChartXml(XSSFChart chart, out XmlNamespaceManager ns)
        {
            XmlDocument xml = new XmlDocument();
            using (Stream stream = chart.GetPackagePart().GetInputStream())
            {
                xml.Load(stream);
            }
            ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", ChartNamespace);
            return xml;
        }

        private static XSSFChart SingleChart(XSSFWorkbook wb)
        {
            XSSFDrawing drawing = ((XSSFSheet)wb.GetSheetAt(0)).GetDrawingPatriarch();
            ClassicAssert.AreEqual(1, drawing.GetCharts().Count);
            return drawing.GetCharts()[0];
        }
    }
}
