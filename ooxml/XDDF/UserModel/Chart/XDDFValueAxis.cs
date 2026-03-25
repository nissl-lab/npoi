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


using System;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XDDF.UserModel.Chart
{
    public class XDDFValueAxis : XDDFChartAxis
    {

        private CT_ValAx ctValAx;

        public XDDFValueAxis(CT_PlotArea plotArea, AxisPosition position)
        {
            initializeAxis(plotArea, position);
        }

        public XDDFValueAxis(CT_ValAx ctValAx)
        {
            this.ctValAx = ctValAx;
        }
        public override XDDFShapeProperties GetOrAddMajorGridProperties()
        {
            CT_ChartLines majorGridlines;
            if(ctValAx.IsSetMajorGridlines())
            {
                majorGridlines = ctValAx.majorGridlines;
            }
            else
            {
                majorGridlines = ctValAx.AddNewMajorGridlines();
            }
            return new XDDFShapeProperties(GetOrAddLinesProperties(majorGridlines));
        }
        public override XDDFShapeProperties GetOrAddMinorGridProperties()
        {
            CT_ChartLines minorGridlines;
            if(ctValAx.IsSetMinorGridlines())
            {
                minorGridlines = ctValAx.minorGridlines;
            }
            else
            {
                minorGridlines = ctValAx.AddNewMinorGridlines();
            }
            return new XDDFShapeProperties(GetOrAddLinesProperties(minorGridlines));
        }
        public override XDDFShapeProperties GetOrAddShapeProperties()
        {
            CT_ShapeProperties properties;
            if(ctValAx.IsSetSpPr())
            {
                properties = ctValAx.spPr;
            }
            else
            {
                properties = ctValAx.AddNewSpPr();
            }
            return new XDDFShapeProperties(properties);
        }
        public override bool IsSetMinorUnit()
        {
            return ctValAx.IsSetMinorUnit();
        }
        public override void SetMinorUnit(double minor)
        {
            if(Double.IsNaN(minor))
            {
                if(ctValAx.IsSetMinorUnit())
                {
                    ctValAx.UnsetMinorUnit();
                }
            }
            else
            {
                if(ctValAx.IsSetMinorUnit())
                {
                    ctValAx.minorUnit.val = minor;
                }
                else
                {
                    ctValAx.AddNewMinorUnit().val = minor;
                }
            }
        }
        public override double GetMinorUnit()
        {
            if(ctValAx.IsSetMinorUnit())
            {
                return ctValAx.minorUnit.val;
            }
            else
            {
                return Double.NaN;
            }
        }
        public override bool IsSetMajorUnit()
        {
            return ctValAx.IsSetMajorUnit();
        }
        public override void SetMajorUnit(double major)
        {
            if(Double.IsNaN(major))
            {
                if(ctValAx.IsSetMajorUnit())
                {
                    ctValAx.UnsetMajorUnit();
                }
            }
            else
            {
                if(ctValAx.IsSetMajorUnit())
                {
                    ctValAx.majorUnit.val = major;
                }
                else
                {
                    ctValAx.AddNewMajorUnit().val = major;
                }
            }
        }
        public override double GetMajorUnit()
        {
            if(ctValAx.IsSetMajorUnit())
            {
                return ctValAx.majorUnit.val;
            }
            else
            {
                return Double.NaN;
            }
        }
        public override void CrossAxis(XDDFChartAxis axis)
        {
            ctValAx.crossAx.val = (uint)axis.Id;
        }
        protected override CT_UnsignedInt GetCTAxId()
        {
            return ctValAx.axId;
        }
        protected override CT_AxPos GetCTAxPos()
        {
            return ctValAx.axPos;
        }
        public override bool HasNumberFormat()
        {
            return ctValAx.IsSetNumFmt();
        }
        protected override CT_NumFmt GetCTNumFmt()
        {
            if(ctValAx.IsSetNumFmt())
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
            CT_Crosses crosses = ctValAx.crosses;
            if(crosses == null)
            {
                return ctValAx.AddNewCrosses();
            }
            else
            {
                return crosses;
            }
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

        public AxisCrossBetween CrossBetween
        {
            get
            {
                return AxisCrossBetweenExtensions.ValueOf(ctValAx.crossBetween.val);
            }
            set 
            {
                ctValAx.crossBetween.val = value.ToST_CrossBetween();
            }
        }

        private void initializeAxis(CT_PlotArea plotArea, AxisPosition position)
        {
            long id = GetNextAxId(plotArea);
            ctValAx = plotArea.AddNewValAx();
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

            Position = (position);
            Orientation = (AxisOrientation.MinMax);
            CrossBetween =(AxisCrossBetween.MidpointCategory);
            Crosses=(AxisCrosses.AutoZero);
            IsVisible=(true);
            MajorTickMark=(AxisTickMark.Cross);
            MinorTickMark=(AxisTickMark.None);
        }

        public override void SetTitle(string text)
        {
            if(!ctValAx.IsSetTitle())
            {
                ctValAx.AddNewTitle();
            }
            XDDFTitle title = new XDDFTitle(null, ctValAx.title);
            title.SetOverlay(false);
            title.SetText(text);
        }
    }
}


