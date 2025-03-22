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
    public class XDDFSeriesAxis : XDDFChartAxis
    {

        private CT_SerAx ctSerAx;

        public XDDFSeriesAxis(CT_PlotArea plotArea, AxisPosition position)
        {
            initializeAxis(plotArea, position);
        }

        public XDDFSeriesAxis(CT_SerAx ctSerAx)
        {
            this.ctSerAx = ctSerAx;
        }
        public override XDDFShapeProperties GetOrAddMajorGridProperties()
        {
            CT_ChartLines majorGridlines;
            if(ctSerAx.IsSetMajorGridlines())
            {
                majorGridlines = ctSerAx.majorGridlines;
            }
            else
            {
                majorGridlines = ctSerAx.AddNewMajorGridlines();
            }
            return new XDDFShapeProperties(GetOrAddLinesProperties(majorGridlines));
        }
        public override XDDFShapeProperties GetOrAddMinorGridProperties()
        {
            CT_ChartLines minorGridlines;
            if(ctSerAx.IsSetMinorGridlines())
            {
                minorGridlines = ctSerAx.minorGridlines;
            }
            else
            {
                minorGridlines = ctSerAx.AddNewMinorGridlines();
            }
            return new XDDFShapeProperties(GetOrAddLinesProperties(minorGridlines));
        }
        public override XDDFShapeProperties GetOrAddShapeProperties()
        {
            OpenXmlFormats.Dml.Chart.CT_ShapeProperties properties;
            if(ctSerAx.IsSetSpPr())
            {
                properties = ctSerAx.spPr;
            }
            else
            {
                properties = ctSerAx.AddNewSpPr();
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
            ctSerAx.crossAx.val = (uint)axis.GetId();
        }
        protected override CT_UnsignedInt GetCTAxId()
        {
            return ctSerAx.axId;
        }
        protected override CT_AxPos GetCTAxPos()
        {
            return ctSerAx.axPos;
        }
        public override bool HasNumberFormat()
        {
            return ctSerAx.IsSetNumFmt();
        }
        protected override CT_NumFmt GetCTNumFmt()
        {
            if(ctSerAx.IsSetNumFmt())
            {
                return ctSerAx.numFmt;
            }
            return ctSerAx.AddNewNumFmt();
        }
        protected override CT_Scaling GetCTScaling()
        {
            return ctSerAx.scaling;
        }
        protected override CT_Crosses GetCTCrosses()
        {
            CT_Crosses crosses = ctSerAx.crosses;
            if(crosses == null)
            {
                return ctSerAx.AddNewCrosses();
            }
            else
            {
                return crosses;
            }
        }
        protected override CT_Boolean GetDelete()
        {
            return ctSerAx.delete;
        }
        protected override CT_TickMark GetMajorCTTickMark()
        {
            return ctSerAx.majorTickMark;
        }
        protected override CT_TickMark GetMinorCTTickMark()
        {
            return ctSerAx.minorTickMark;
        }

        private void initializeAxis(CT_PlotArea plotArea, AxisPosition position)
        {
            long id = GetNextAxId(plotArea);
            ctSerAx = plotArea.AddNewSerAx();
            ctSerAx.AddNewAxId().val = (uint)id;
            ctSerAx.AddNewAxPos();
            ctSerAx.AddNewScaling();
            ctSerAx.AddNewCrosses();
            ctSerAx.AddNewCrossAx();
            ctSerAx.AddNewTickLblPos().val = ST_TickLblPos.nextTo;
            ctSerAx.AddNewDelete();
            ctSerAx.AddNewMajorTickMark();
            ctSerAx.AddNewMinorTickMark();

            SetPosition(position);
            SetOrientation(AxisOrientation.MinMax);
            SetCrosses(AxisCrosses.AutoZero);
            SetVisible(true);
            SetMajorTickMark(AxisTickMark.Cross);
            SetMinorTickMark(AxisTickMark.None);
        }
    }
}


