using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace NPOI.XSSF.UserModel.Charts
{
    /// <summary>
    /// Holds data for a XSSF Bar Chart
    /// </summary>
    /// <typeparam name="Tx"></typeparam>
    /// <typeparam name="Ty"></typeparam>
    public class XSSFBarChartData<Tx, Ty> : IBarChartData<Tx, Ty>
    {
        /**
         * List of all data series.
         */
        private List<IBarChartSeries<Tx, Ty>> series;

        public XSSFBarChartData()
        {
            series = new List<IBarChartSeries<Tx, Ty>>();
        }

        public class Series : AbstractXSSFChartSeries, IBarChartSeries<Tx, Ty>
        {
            private int id;
            private int order;
            private byte[] fillColor;
            private IChartDataSource<Tx> categories;
            private IChartDataSource<Ty> values;

            internal Series(int id, int order,
                IChartDataSource<Tx> categories,
                IChartDataSource<Ty> values)
            {
                this.id = id;
                this.order = order;
                this.categories = categories;
                this.values = values;
            }

            public void SetId(int id)
            {
                this.id = id;
            }

            public void SetOrder(int order)
            {
                this.order = order;
            }

            public void SetFillColor(Color color)
            {
                fillColor = new byte[3];
                fillColor[0] = color.R;
                fillColor[1] = color.G;
                fillColor[2] = color.B;
            }

            public IChartDataSource<Tx> GetCategoryAxisData()
            {
                return categories;
            }

            public IChartDataSource<Ty> GetValues()
            {
                return values;
            }

            internal void AddToChart(CT_BarChart ctBarChart)
            {
                CT_BarSer ctBarSer = ctBarChart.AddNewSer();
                CT_BarGrouping ctGrouping = ctBarChart.AddNewGrouping();
                ctGrouping.val = ST_BarGrouping.clustered;
                ctBarSer.AddNewIdx().val = (uint)id;
                ctBarSer.AddNewOrder().val = (uint)order;
                CT_Boolean ctNoInvertIfNegative = new CT_Boolean();
                ctNoInvertIfNegative.val = 0;
                ctBarSer.invertIfNegative = ctNoInvertIfNegative;

                CT_BarDir ctBarDir = ctBarChart.AddNewBarDir();
                ctBarDir.val = ST_BarDir.bar;

                CT_AxDataSource catDS = ctBarSer.AddNewCat();
                XSSFChartUtil.BuildAxDataSource(catDS, categories);
                CT_NumDataSource valueDS = ctBarSer.AddNewVal();
                XSSFChartUtil.BuildNumDataSource(valueDS, values);

                if (IsTitleSet)
                {
                    ctBarSer.tx = GetCTSerTx();
                }

                if (fillColor != null)
                {
                    ctBarSer.spPr = new OpenXmlFormats.Dml.Chart.CT_ShapeProperties();
                    CT_SolidColorFillProperties ctSolidColorFillProperties = ctBarSer.spPr.AddNewSolidFill();
                    CT_SRgbColor ctSRgbColor = ctSolidColorFillProperties.AddNewSrgbClr();
                    ctSRgbColor.val = fillColor;
                }
            }
        }

        public IBarChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categoryAxisData, IChartDataSource<Ty> values)
        {
            if (!values.IsNumeric)
            {
                throw new ArgumentException("Value data source must be numeric.");
            }
            int numOfSeries = series.Count;
            Series newSeries = new Series(numOfSeries, numOfSeries, categoryAxisData, values);
            series.Add(newSeries);
            return newSeries;
        }

        public List<IBarChartSeries<Tx, Ty>> GetSeries()
        {
            return series;
        }

        public void FillChart(SS.UserModel.IChart chart, params IChartAxis[] axis)
        {
            if (!(chart is XSSFChart))
            {
                throw new ArgumentException("Chart must be instance of XSSFChart");
            }

            XSSFChart xssfChart = (XSSFChart)chart;
            CT_PlotArea plotArea = xssfChart.GetCTChart().plotArea;
            int allSeriesCount = plotArea.GetAllSeriesCount();
            CT_BarChart barChart = plotArea.AddNewBarChart();
            barChart.AddNewVaryColors().val = 0;

            for (int i = 0; i < series.Count; ++i)
            {
                Series s = (Series)series[i];
                s.SetId(allSeriesCount + i);
                s.SetOrder(allSeriesCount + i);
                s.AddToChart(barChart);
            }

            foreach (IChartAxis ax in axis)
            {
                barChart.AddNewAxId().val = (uint)ax.Id;
            }
        }
    }
}
