
namespace NPOI.XDDF.UserModel.Chart;

using System.Collections.Generic;
using NPOI.OpenXmlFormats.Dml.Chart;

public class XDDFArea3DChartData<T, V> : XDDFChartData<T, V>
{
    private CT_Area3DChart chart;

    public XDDFArea3DChartData(CT_Area3DChart chart, Dictionary<long, XDDFChartAxis> categories,
            Dictionary<long, XDDFValueAxis> values)
    {
        this.chart = chart;
        if (chart.ser != null)
        {
            foreach (CT_AreaSer series in chart.ser)
            {
                this.series.Add(new Series(series, series.cat, series.val));
            }
        }
        DefineAxes(categories, values);
    }

    private void DefineAxes(Dictionary<long, XDDFChartAxis> categories, Dictionary<long, XDDFValueAxis> values)
    {
        if (chart.axId == null || chart.axId.Count == 0)
        {
            if (chart.axId == null) chart.axId = [];
            foreach (long id in categories.Keys)
            {
                chart.axId.Add(new CT_UnsignedInt { val = (uint)id });
            }
            foreach (long id in values.Keys)
            {
                chart.axId.Add(new CT_UnsignedInt { val = (uint)id });
            }
        }
        DefineAxis([.. chart.axId], categories, values);
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

    public Grouping GetGrouping()
    {
        return GroupingExtensions.ValueOf(chart.grouping.val);
    }

    public void SetGrouping(Grouping grouping)
    {
        if (chart.grouping != null)
        {
            chart.grouping.val = grouping.ToST_Grouping();
        }
        else
        {
            chart.grouping = new CT_Grouping { val = grouping.ToST_Grouping() };
        }
    }

    public override XDDFChartData<T, V>.Series AddSeries(IXDDFDataSource<T> category,
            IXDDFNumericalDataSource<V> values)
    {
        int index = this.series.Count;
        CT_AreaSer ctSer = new CT_AreaSer();
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
        private CT_AreaSer series;

        internal Series(CT_AreaSer series, IXDDFDataSource<T> category,
                IXDDFNumericalDataSource<V> values)
            : base(category, values)
        {
            this.series = series;
        }

        internal Series(CT_AreaSer series, CT_AxDataSource category, CT_NumDataSource values)
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
                    series.spPr = new NPOI.OpenXmlFormats.Dml.CT_ShapeProperties();
                    series.spPr.Set(properties.GetXmlObject());
                }
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
