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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.Util;
    using NPOI.XDDF.UserModel;
    using NPOI.OpenXmlFormats.Dml.Chart;
    using NPOI.OpenXmlFormats.Dml;
    public class XDDFCategoryAxis : XDDFChartAxis
    {

        private CT_CatAx ctCatAx;

        public XDDFCategoryAxis(CT_PlotArea plotArea, AxisPosition position)
        {
            initializeAxis(plotArea, position);
        }

        public XDDFCategoryAxis(CT_CatAx ctCatAx)
        {
            this.ctCatAx = ctCatAx;
        }
        public override XDDFShapeProperties GetOrAddMajorGridProperties()
        {
            CT_ChartLines majorGridlines;
            if(ctCatAx.IsSetMajorGridlines())
            {
                majorGridlines = ctCatAx.majorGridlines;
            }
            else
            {
                majorGridlines = ctCatAx.AddNewMajorGridlines();
            }
            return new XDDFShapeProperties(GetOrAddLinesProperties(majorGridlines));
        }
        public override XDDFShapeProperties GetOrAddMinorGridProperties()
        {
            CT_ChartLines minorGridlines;
            if(ctCatAx.IsSetMinorGridlines())
            {
                minorGridlines = ctCatAx.minorGridlines;
            }
            else
            {
                minorGridlines = ctCatAx.AddNewMinorGridlines();
            }
            return new XDDFShapeProperties(GetOrAddLinesProperties(minorGridlines));
        }
        public override XDDFShapeProperties GetOrAddShapeProperties()
        {
            OpenXmlFormats.Dml.Chart.CT_ShapeProperties properties;
            if(ctCatAx.IsSetSpPr())
            {
                properties = ctCatAx.spPr;
            }
            else
            {
                properties = ctCatAx.AddNewSpPr();
            }

            return new XDDFShapeProperties(properties);
        }
        public override bool IsSetMinorUnit()
        {
            return false;
        }
        public override void SetMinorUnit(double minor)
        {
            // nothing
        }
        public override double GetMinorUnit()
        {
            return Double.NaN;
        }
        public override bool IsSetMajorUnit()
        {
            return false;
        }
        public override void SetMajorUnit(double major)
        {
            // nothing
        }
        public override double GetMajorUnit()
        {
            return Double.NaN;
        }
        public override void crossAxis(XDDFChartAxis axis)
        {
            ctCatAx.crossAx.val = (uint)axis.GetId();
        }
        protected override CT_UnsignedInt GetCTAxId()
        {
            return ctCatAx.axId;
        }
        protected override CT_AxPos GetCTAxPos()
        {
            return ctCatAx.axPos;
        }
        public override bool HasNumberFormat()
        {
            return ctCatAx.IsSetNumFmt();
        }
        protected override CT_NumFmt GetCTNumFmt()
        {
            if(ctCatAx.IsSetNumFmt())
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
            CT_Crosses crosses = ctCatAx.crosses;
            if(crosses == null)
            {
                return ctCatAx.AddNewCrosses();
            }
            else
            {
                return crosses;
            }
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

        public AxisLabelAlignment GetLabelAlignment()
        {
            return AxisLabelAlignmentExtensions.ValueOf(ctCatAx.lblAlgn.val);
        }

        public void SetLabelAlignment(AxisLabelAlignment labelAlignment)
        {
            ctCatAx.lblAlgn.val = labelAlignment.ToST_LblAlgn();
        }

        private void initializeAxis(CT_PlotArea plotArea, AxisPosition position)
        {
            long id = GetNextAxId(plotArea);
            ctCatAx = plotArea.AddNewCatAx();
            ctCatAx.AddNewAxId().val = (uint)id;
            ctCatAx.AddNewAxPos();
            ctCatAx.AddNewScaling();
            ctCatAx.AddNewCrosses();
            ctCatAx.AddNewCrossAx();
            ctCatAx.AddNewTickLblPos().val = ST_TickLblPos.nextTo;
            ctCatAx.AddNewDelete();
            ctCatAx.AddNewMajorTickMark();
            ctCatAx.AddNewMinorTickMark();

            SetPosition(position);
            SetOrientation(AxisOrientation.MinMax);
            SetCrosses(AxisCrosses.AutoZero);
            SetVisible(true);
            SetMajorTickMark(AxisTickMark.Cross);
            SetMinorTickMark(AxisTickMark.None);
        }
    }
}


