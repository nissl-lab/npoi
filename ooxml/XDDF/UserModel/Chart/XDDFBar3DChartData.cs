/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.XDDF.UserModel.Chart;

using System.Collections.Generic;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Chart;

public class XDDFBar3DChartData<T, V> : XDDFChartData<T, V>
{
    private CT_Bar3DChart chart;

    public XDDFBar3DChartData(CT_Bar3DChart chart, Dictionary<long, XDDFChartAxis> categories,
            Dictionary<long, XDDFValueAxis> values)
    {
        this.chart = chart;
        if (chart.barDir == null)
        {
            chart.barDir = new CT_BarDir { val = BarDirection.Bar.ToST_BarDir() };
        }
        if (chart.ser != null)
        {
            foreach (CT_BarSer series in chart.ser)
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

    public BarDirection GetBarDirection()
    {
        return BarDirectionExtensions.ValueOf(chart.barDir.val);
    }

    public void SetBarDirection(BarDirection direction)
    {
        chart.barDir.val = direction.ToST_BarDir();
    }

    public BarGrouping GetBarGrouping()
    {
        if (chart.grouping != null)
        {
            return BarGroupingExtensions.ValueOf(chart.grouping.val);
        }
        else
        {
            return BarGrouping.Standard;
        }
    }

    public void SetBarGrouping(BarGrouping grouping)
    {
        if (chart.grouping != null)
        {
            chart.grouping.val = grouping.ToST_BarGrouping();
        }
        else
        {
            chart.grouping = new CT_BarGrouping { val = grouping.ToST_BarGrouping() };
        }
    }

    public int GetGapWidth()
    {
        if (chart.gapWidth != null)
        {
            return chart.gapWidth.val;
        }
        else
        {
            return 0;
        }
    }

    public void SetGapWidth(int width)
    {
        if (chart.gapWidth != null)
        {
            chart.gapWidth.val = (ushort)width;
        }
        else
        {
            chart.gapWidth = new CT_GapAmount { val = (ushort)width };
        }
    }

    public override XDDFChartData<T, V>.Series AddSeries(IXDDFDataSource<T> category,
            IXDDFNumericalDataSource<V> values)
    {
        int index = this.series.Count;
        CT_BarSer ctSer = new CT_BarSer();
        chart.ser.Add(ctSer);
        ctSer.tx = new CT_SerTx();
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
        private CT_BarSer series;

        internal Series(CT_BarSer series, IXDDFDataSource<T> category,
                IXDDFNumericalDataSource<V> values)
            : base(category, values)
        {
            this.series = series;
        }

        internal Series(CT_BarSer series, CT_AxDataSource category, CT_NumDataSource values)
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
