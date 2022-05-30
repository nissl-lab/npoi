using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;

namespace NPOI.XSSF.UserModel.Charts
{
    /// <summary>
    /// Holds data for a XSSF Pie Chart
    /// </summary>
    /// <typeparam name="Tx"></typeparam>
    /// <typeparam name="Ty"></typeparam>
    public class XSSFPieChartData<Tx, Ty> : IPieChartData<Tx, Ty>
    {
        /// <summary>
        /// List of all data series.
        /// </summary>
        private List<IPieChartSeries<Tx, Ty>> series;

        public XSSFPieChartData()
        {
            series = new List<IPieChartSeries<Tx, Ty>>();
        }

        public class Series : AbstractXSSFChartSeries, IPieChartSeries<Tx, Ty>
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

            public IChartDataSource<Tx> GetCategoryAxisData()
            {
                return categories;
            }

            public IChartDataSource<Ty> GetValues()
            {
                return values;
            }

            internal void AddToChart(CT_PieChart ctPieChart)
            {
                CT_PieSer ctPieSer = ctPieChart.AddNewSer();
                ctPieSer.AddNewIdx().val = (uint) id;
                ctPieSer.AddNewOrder().val = (uint) order;

                CT_AxDataSource catDS = ctPieSer.AddNewCat();
                XSSFChartUtil.BuildAxDataSource(catDS, categories);
                for (var i = 0; i < categories.PointCount; i++)
                {
                    var accentIndex = (i % 6) + 4;
                    var colorVal = (ST_SchemeColorVal)Enum.Parse(typeof(ST_SchemeColorVal), accentIndex.ToString());
                    var schemeColor = new CT_SchemeColor {val = colorVal};
                    schemeColor.AddNewLum(i);
                    var solidFill = new CT_SolidColorFillProperties {schemeClr = schemeColor};
                    var shapeProperties = new NPOI.OpenXmlFormats.Dml.Chart.CT_ShapeProperties {solidFill = solidFill};
                    var dPt = new CT_DPt
                    {
                        spPr = shapeProperties,
                        idx = new CT_UnsignedInt {val = (uint) i},
                    };
                    
                    ctPieSer.dPt.Add(dPt);
                }

                CT_NumDataSource valueDS = ctPieSer.AddNewVal();
                XSSFChartUtil.BuildNumDataSource(valueDS, values);

                if (IsTitleSet)
                {
                    ctPieSer.tx = GetCTSerTx();
                }
            }
        }

        public IPieChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categoryAxisData, IChartDataSource<Ty> values)
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

        public List<IPieChartSeries<Tx, Ty>> GetSeries()
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

            CT_PieChart pieChart = plotArea.AddNewPieChart();
            pieChart.AddNewVaryColors().val = 0;

            for (int i = 0; i < series.Count; ++i)
            {
                Series s = (Series)series[i];
                s.SetId(allSeriesCount + i);
                s.SetOrder(allSeriesCount + i);
                s.AddToChart(pieChart);
            }
        }
    }
}
