using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel.Charts;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.XSSF.UserModel.Charts
{

    public class XSSFCategoryAxis : XSSFChartAxis
    {

        private CT_CatAx ctCatAx;

        public XSSFCategoryAxis(XSSFChart chart, long id, AxisPosition pos)
            : base(chart)
        {

            createAxis(id, pos);
        }

        public XSSFCategoryAxis(XSSFChart chart, CT_CatAx ctCatAx)
            : base(chart)
        {

            this.ctCatAx = ctCatAx;
        }

        public override long Id
        {
            get
            {
                return ctCatAx.axId.val;
            }
        }

        protected override CT_AxPos GetCTAxPos()
        {
            return ctCatAx.axPos;
        }

        protected override CT_NumFmt GetCTNumFmt()
        {
            if (ctCatAx.IsSetNumFmt())
            {
                return ctCatAx.numFmt;
            }
            return ctCatAx.AddNewNumFmt();
        }

        protected override CT_Scaling GetCTScaling()
        {
            return ctCatAx.scaling;
        }

        protected override CT_Crosses GetCTCrosses()
        {
            return ctCatAx.crosses;
        }

        protected override CT_Boolean GetDelete()
        {
            return ctCatAx.delete;
        }

        protected override CT_TickMark GetMajorCTTickMark()
        {
            return ctCatAx.majorTickMark;
        }

        protected override CT_TickMark GetMinorCTTickMark()
        {
            return ctCatAx.minorTickMark;
        }

        public override void CrossAxis(IChartAxis axis)
        {
            ctCatAx.crossAx.val = (uint)axis.Id;
        }

        private void createAxis(long id, AxisPosition pos)
        {
            ctCatAx = chart.GetCTChart().plotArea.AddNewCatAx();
            ctCatAx.AddNewAxId().val = (uint)id;
            ctCatAx.AddNewAxPos();
            ctCatAx.AddNewScaling();
            ctCatAx.AddNewCrosses();
            ctCatAx.AddNewCrossAx();
            ctCatAx.AddNewTickLblPos().val = ST_TickLblPos.nextTo;
            ctCatAx.AddNewDelete();
            ctCatAx.AddNewMajorTickMark();
            ctCatAx.AddNewMinorTickMark();

            
            this.Position=(pos);
            this.Orientation=(AxisOrientation.MinToMax);
            this.Crosses=(AxisCrosses.AutoZero);
            this.IsVisible = true;
            this.MajorTickMark=(AxisTickMark.Cross);
            this.MinorTickMark=(AxisTickMark.None);
        }
    }
}

