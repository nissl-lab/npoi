using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.XSSF.UserModel.Charts
{
    public class XSSFDateAxis : XSSFChartAxis
    {
        private CT_DateAx ctDateAx;

        public XSSFDateAxis(XSSFChart chart, long id, AxisPosition pos)
            : base(chart)
        {

            createAxis(id, pos);
        }

        public XSSFDateAxis(XSSFChart chart, CT_DateAx ctDateAx)
            : base(chart)
        {

            this.ctDateAx = ctDateAx;
        }

        public override long Id
        {
            get
            {
                return ctDateAx.axId.val;
            }
        }

        public CT_ShapeProperties Line
        {
            get
            {
                return ctDateAx.spPr;
            }
        }

        protected override CT_AxPos GetCTAxPos()
        {
            return ctDateAx.axPos;
        }

        protected override CT_NumFmt GetCTNumFmt()
        {
            if (ctDateAx.IsSetNumFmt())
            {
                return ctDateAx.numFmt;
            }
            return ctDateAx.AddNewNumFmt();
        }

        protected override CT_Scaling GetCTScaling()
        {
            return ctDateAx.scaling;
        }

        protected override CT_Crosses GetCTCrosses()
        {
            return ctDateAx.crosses;
        }

        protected override CT_Boolean GetDelete()
        {
            return ctDateAx.delete;
        }

        protected override CT_TickMark GetMajorCTTickMark()
        {
            return ctDateAx.majorTickMark;
        }

        public void SetMajorCTTickMark(CT_TickMark tm)
        {
            ctDateAx.majorTickMark = tm;
        }

        protected override CT_TickMark GetMinorCTTickMark()
        {
            return ctDateAx.minorTickMark;
        }

        protected CT_ChartLines GetMajorGridLines()
        {
            return ctDateAx.majorGridlines;
        }

        public override void CrossAxis(IChartAxis axis)
        {
            ctDateAx.crossAx.val = (uint)axis.Id;
        }

        public CT_TimeUnit GetBaseTimeUnit()
        {
            return ctDateAx.baseTimeUnit;
        }

        public void SetBaseTimeUnit(CT_TimeUnit unit)
        {
            ctDateAx.baseTimeUnit = unit;
        }

        public void SetAuto(CT_Boolean au)
        {
            ctDateAx.auto = au;
        }

        private void createAxis(long id, AxisPosition pos)
        {
            ctDateAx = chart.GetCTChart().plotArea.AddNewDateAx();
            ctDateAx.AddNewAxId().val = (uint)id;
            ctDateAx.AddNewAxPos();
            ctDateAx.AddNewScaling();
            ctDateAx.AddNewCrosses();
            ctDateAx.AddNewCrossAx();
            ctDateAx.AddNewTickLblPos().val = ST_TickLblPos.nextTo;
            ctDateAx.AddNewDelete();
            ctDateAx.AddNewMajorTickMark();
            ctDateAx.AddNewMinorTickMark();


            this.Position = (pos);
            this.Orientation = (AxisOrientation.MinToMax);
            this.Crosses = (AxisCrosses.AutoZero);
            this.IsVisible = true;
            this.MajorTickMark = (AxisTickMark.Cross);
            this.MinorTickMark = (AxisTickMark.None);
        }
    }
}
