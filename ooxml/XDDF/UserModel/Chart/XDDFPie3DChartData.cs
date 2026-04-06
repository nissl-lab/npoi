
namespace NPOI.XDDF.UserModel.Chart;

using System.Collections.Generic;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Chart;

public class XDDFPie3DChartData<T, V> : XDDFChartData<T, V>
{
    private CT_Pie3DChart chart;

    public XDDFPie3DChartData(CT_Pie3DChart chart)
    {
        this.chart = chart;
        if (chart.ser != null)
        {
            foreach (CT_PieSer series in chart.ser)
            {
                this.series.Add(new Series(series, series.cat, series.val));
            }
        }
    }

    public override void SetVaryColors(bool varyColors)
    {
        if (chart.varyColors != null)
        {
            chart.varyColors.val = varyColors ? 1 : 0;
        }
        else
        {
            chart.varyColors = new CT_Boolean { val = varyColors ? 1 : 0 };
        }
    }

    public override XDDFChartData<T, V>.Series AddSeries(IXDDFDataSource<T> category,
            IXDDFNumericalDataSource<V> values)
    {
        int index = this.series.Count;
        CT_PieSer ctSer = new CT_PieSer();
        chart.ser.Add(ctSer);
        ctSer.cat = new CT_AxDataSource();
        ctSer.val = new CT_NumDataSource();
        ctSer.idx = new CT_UnsignedInt { val = (uint)index };
        ctSer.order = new CT_UnsignedInt { val = (uint)index };
        Series added = new Series(ctSer, category, values);
        this.series.Add(added);
        return added;
    }

    public class Series : XDDFChartData<T, V>.Series
    {
        private CT_PieSer series;

        internal Series(CT_PieSer series, IXDDFDataSource<T> category,
                IXDDFNumericalDataSource<V> values)
            : base(category, values)
        {
            this.series = series;
        }

        internal Series(CT_PieSer series, CT_AxDataSource category, CT_NumDataSource values)
            : base(XDDFDataSourcesFactory.FromDataSource(category) as IXDDFDataSource<T>,
                  XDDFDataSourcesFactory.FromDataSource(values) as IXDDFNumericalDataSource<V>)
        {
            this.series = series;
        }

        protected override CT_SerTx GetSeriesText()
        {
            if (series.tx != null)
            {
                return series.tx;
            }
            else
            {
                series.tx = new CT_SerTx();
                return series.tx;
            }
        }

        public override void SetShowLeaderLines(bool showLeaderLines)
        {
            if (series.dLbls == null)
            {
                series.dLbls = new CT_DLbls();
            }
            if (series.dLbls.showLeaderLines != null)
            {
                series.dLbls.showLeaderLines.val = showLeaderLines ? 1 : 0;
            }
            else
            {
                series.dLbls.showLeaderLines = new CT_Boolean { val = showLeaderLines ? 1 : 0 };
            }
        }

        public override XDDFShapeProperties GetShapeProperties()
        {
            if (series.spPr != null)
            {
                return new XDDFShapeProperties(series.spPr);
            }
            else
            {
                return null;
            }
        }

        public override void SetShapeProperties(XDDFShapeProperties properties)
        {
            if (properties == null)
            {
                if (series.spPr != null)
                {
                    series.spPr = null;
                }
            }
            else
            {
                if (series.spPr != null)
                {
                    series.spPr = properties.GetXmlObject();
                }
                else
                {
                    series.spPr = new CT_ShapeProperties();
                    series.spPr.Set(properties.GetXmlObject());
                }
            }
        }

        public long GetExplosion()
        {
            if (series.explosion != null)
            {
                return series.explosion.val;
            }
            else
            {
                return 0;
            }
        }

        public void SetExplosion(long explosion)
        {
            if (series.explosion != null)
            {
                series.explosion.val = (uint)explosion;
            }
            else
            {
                series.explosion = new CT_UnsignedInt { val = (uint)explosion };
            }
        }

        protected override CT_AxDataSource GetAxDS()
        {
            return series.cat;
        }

        protected override CT_NumDataSource GetNumDS()
        {
            return series.val;
        }

        public void UpdateIdXVal(long val)
        {
            series.idx.val = (uint)val;
        }

        public void UpdateOrderVal(long val)
        {
            series.order.val = (uint)val;
        }
    }
}
