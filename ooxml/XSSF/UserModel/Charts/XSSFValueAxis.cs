/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using NPOI.SS.UserModel.Charts;
using NPOI.OpenXmlFormats.Dml.Chart;
using System;
namespace NPOI.XSSF.UserModel.Charts
{

    /**
     * Value axis type.
     *
     * @author Roman Kashitsyn
     */
    public class XSSFValueAxis : XSSFChartAxis, IValueAxis
    {

        private CT_ValAx ctValAx;

        public XSSFValueAxis(XSSFChart chart, long id, AxisPosition pos)
            : base(chart)
        {

            CreateAxis(id, pos);
        }

        public XSSFValueAxis(XSSFChart chart, CT_ValAx ctValAx)
            : base(chart)
        {

            this.ctValAx = ctValAx;
        }

        public override long Id
        {
            get
            {
                return ctValAx.axId.val;
            }
        }

        public void SetCrossBetween(AxisCrossBetween crossBetween)
        {
            ctValAx.crossBetween.val= fromCrossBetween(crossBetween);
        }

        public AxisCrossBetween GetCrossBetween()
        {
            return ToCrossBetween(ctValAx.crossBetween.val);
        }

        protected override CT_Boolean GetDelete()
        {
            return ctValAx.delete;
        }

        protected override CT_TickMark GetMajorCTTickMark()
        {
            return ctValAx.majorTickMark;
        }

        protected override CT_TickMark GetMinorCTTickMark()
        {
            return ctValAx.minorTickMark;
        }
        protected override CT_AxPos GetCTAxPos()
        {
            return ctValAx.axPos;
        }


        protected override CT_NumFmt GetCTNumFmt()
        {
            if (ctValAx.IsSetNumFmt())
            {
                return ctValAx.numFmt;
            }
            return ctValAx.AddNewNumFmt();
        }
        protected override CT_Scaling GetCTScaling()
        {
            return ctValAx.scaling;
        }


        protected override CT_Crosses GetCTCrosses()
        {
            return ctValAx.crosses;
        }

        public override void CrossAxis(IChartAxis axis)
        {
            ctValAx.crossAx.val= (uint)axis.Id;
        }

        private void CreateAxis(long id, AxisPosition pos)
        {
            ctValAx = chart.GetCTChart().plotArea.AddNewValAx();
            ctValAx.AddNewAxId().val = (uint)id;
            ctValAx.AddNewAxPos();
            ctValAx.AddNewScaling();
            ctValAx.AddNewCrossBetween();
            ctValAx.AddNewCrosses();
            ctValAx.AddNewCrossAx();
            ctValAx.AddNewTickLblPos().val = ST_TickLblPos.nextTo;
            ctValAx.AddNewDelete();
            ctValAx.AddNewMajorTickMark();
            ctValAx.AddNewMinorTickMark();

            Position=(pos);
            Orientation=(AxisOrientation.MinToMax);
            SetCrossBetween(AxisCrossBetween.MidpointCategory);
            Crosses=(AxisCrosses.AutoZero);
            IsVisible=(true);
            MajorTickMark = (AxisTickMark.Cross);
            MinorTickMark=(AxisTickMark.None);
        }

        private static ST_CrossBetween fromCrossBetween(AxisCrossBetween crossBetween)
        {
            switch (crossBetween)
            {
                case AxisCrossBetween.Between: return ST_CrossBetween.between;
                case AxisCrossBetween.MidpointCategory: return ST_CrossBetween.midCat;
                default:
                    throw new ArgumentException();
            }
        }

        private static AxisCrossBetween ToCrossBetween(ST_CrossBetween ctCrossBetween)
        {
            switch (ctCrossBetween)
            {
                case ST_CrossBetween.between: return AxisCrossBetween.Between;
                case ST_CrossBetween.midCat: return AxisCrossBetween.MidpointCategory;
                default:
                    throw new ArgumentException();
            }
        }
    }
}
