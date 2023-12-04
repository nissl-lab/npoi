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
    /// Holds data for a XSSF Area Chart
    /// </summary>
    /// <typeparam name="Tx"></typeparam>
    /// <typeparam name="Ty"></typeparam>
    public class XSSFAreaChartData<Tx, Ty> : IAreaChartData<Tx, Ty>
    {
        /**
         * List of all data series.
         */
        private List<IAreaChartSeries<Tx, Ty>> series;

        public XSSFAreaChartData()
        {
            series = new List<IAreaChartSeries<Tx, Ty>>();
        }

        public class Series : AbstractXSSFChartSeries, IAreaChartSeries<Tx, Ty>
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

            internal void AddToChart(CT_AreaChart ctAreaChart)
            {
                CT_AreaSer ctAreaSer = ctAreaChart.AddNewSer();
                //TODO: resolve grouping
                //CT_BarGrouping ctGrouping = ctAreaSer.AddNewGrouping();
                //ctGrouping.val = ST_BarGrouping.clustered;
                ctAreaSer.AddNewIdx().val = (uint)id;
                ctAreaSer.AddNewOrder().val = (uint)order;
                //TODO: resolve negative
                //CT_Boolean ctNoInvertIfNegative = new CT_Boolean();
                //ctNoInvertIfNegative.val = 0;
                //ctAreaSer.invertIfNegative = ctNoInvertIfNegative;

                CT_AxDataSource catDS = ctAreaSer.AddNewCat();
                XSSFChartUtil.BuildAxDataSource(catDS, categories);
                CT_NumDataSource valueDS = ctAreaSer.AddNewVal();
                XSSFChartUtil.BuildNumDataSource(valueDS, values);

                if (IsTitleSet)
                {
                    ctAreaSer.tx = GetCTSerTx();
                }

                if (fillColor != null)
                {
                    ctAreaSer.spPr = new NPOI.OpenXmlFormats.Dml.Chart.CT_ShapeProperties();
                    CT_SolidColorFillProperties ctSolidColorFillProperties = ctAreaSer.spPr.AddNewSolidFill();
                    CT_SRgbColor ctSRgbColor = ctSolidColorFillProperties.AddNewSrgbClr();
                    ctSRgbColor.val = fillColor;
                }
            }
        }

        public IAreaChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categoryAxisData, IChartDataSource<Ty> values)
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

        public List<IAreaChartSeries<Tx, Ty>> GetSeries()
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
            CT_AreaChart areaChart = plotArea.AddNewAreaChart();
            areaChart.AddNewVaryColors().val = 0;

            for (int i = 0; i < series.Count; ++i)
            {
                Series s = (Series)series[i];
                s.SetId(allSeriesCount + i);
                s.SetOrder(allSeriesCount + i);
                s.AddToChart(areaChart);
            }

            foreach (IChartAxis ax in axis)
            {
                areaChart.AddNewAxId().val = (uint)ax.Id;
            }
        }
    }
}
