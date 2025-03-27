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

    /// <summary>
    /// Date axis type. Currently only : the same values as
    /// <see cref="XDDFCategoryAxis"/>, since the two are nearly identical.
    /// </summary>
    public class XDDFDateAxis : XDDFChartAxis
    {

        private CT_DateAx ctDateAx;

        public XDDFDateAxis(CT_PlotArea plotArea, AxisPosition position)
        {
            initializeAxis(plotArea, position);
        }

        public XDDFDateAxis(CT_DateAx ctDateAx)
        {
            this.ctDateAx = ctDateAx;
        }
        public override XDDFShapeProperties GetOrAddMajorGridProperties()
        {
            CT_ChartLines majorGridlines;
            if(ctDateAx.IsSetMajorGridlines())
            {
                majorGridlines = ctDateAx.majorGridlines;
            }
            else
            {
                majorGridlines = ctDateAx.AddNewMajorGridlines();
            }
            return new XDDFShapeProperties(GetOrAddLinesProperties(majorGridlines));
        }
        public override XDDFShapeProperties GetOrAddMinorGridProperties()
        {
            CT_ChartLines minorGridlines;
            if(ctDateAx.IsSetMinorGridlines())
            {
                minorGridlines = ctDateAx.minorGridlines;
            }
            else
            {
                minorGridlines = ctDateAx.AddNewMinorGridlines();
            }
            return new XDDFShapeProperties(GetOrAddLinesProperties(minorGridlines));
        }
        public override XDDFShapeProperties GetOrAddShapeProperties()
        {
            OpenXmlFormats.Dml.Chart.CT_ShapeProperties properties;
            if(ctDateAx.IsSetSpPr())
            {
                properties = ctDateAx.spPr;
            }
            else
            {
                properties = ctDateAx.AddNewSpPr();
            }
            return new XDDFShapeProperties(properties);
        }
        public override bool IsSetMinorUnit()
        {
            return ctDateAx.IsSetMinorUnit();
        }
        public override void SetMinorUnit(double minor)
        {
            if(Double.IsNaN(minor))
            {
                if(ctDateAx.IsSetMinorUnit())
                {
                    ctDateAx.UnsetMinorUnit();
                }
            }
            else
            {
                if(ctDateAx.IsSetMinorUnit())
                {
                    ctDateAx.minorUnit.val = minor;
                }
                else
                {
                    ctDateAx.AddNewMinorUnit().val = minor;
                }
            }
        }
        public override double GetMinorUnit()
        {
            if(ctDateAx.IsSetMinorUnit())
            {
                return ctDateAx.minorUnit.val;
            }
            else
            {
                return Double.NaN;
            }
        }
        public override bool IsSetMajorUnit()
        {
            return ctDateAx.IsSetMajorUnit();
        }
        public override void SetMajorUnit(double major)
        {
            if(Double.IsNaN(major))
            {
                if(ctDateAx.IsSetMajorUnit())
                {
                    ctDateAx.UnsetMajorUnit();
                }
            }
            else
            {
                if(ctDateAx.IsSetMajorUnit())
                {
                    ctDateAx.majorUnit.val = major;
                }
                else
                {
                    ctDateAx.AddNewMajorUnit().val = major;
                }
            }
        }
        public override double GetMajorUnit()
        {
            if(ctDateAx.IsSetMajorUnit())
            {
                return ctDateAx.majorUnit.val;
            }
            else
            {
                return Double.NaN;
            }
        }
        public override void crossAxis(XDDFChartAxis axis)
        {
            ctDateAx.crossAx.val = (uint)axis.GetId();
        }
        protected override CT_UnsignedInt GetCTAxId()
        {
            return ctDateAx.axId;
        }
        protected override CT_AxPos GetCTAxPos()
        {
            return ctDateAx.axPos;
        }
        public override bool HasNumberFormat()
        {
            return ctDateAx.IsSetNumFmt();
        }
        protected override CT_NumFmt GetCTNumFmt()
        {
            if(ctDateAx.IsSetNumFmt())
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
            CT_Crosses crosses = ctDateAx.crosses;
            if(crosses == null)
            {
                return ctDateAx.AddNewCrosses();
            }
            else
            {
                return crosses;
            }
        }
        protected override CT_Boolean GetDelete()
        {
            return ctDateAx.delete;
        }
        protected override CT_TickMark GetMajorCTTickMark()
        {
            return ctDateAx.majorTickMark;
        }
        protected override CT_TickMark GetMinorCTTickMark()
        {
            return ctDateAx.minorTickMark;
        }

        private void initializeAxis(CT_PlotArea plotArea, AxisPosition position)
        {
            long id = GetNextAxId(plotArea);
            ctDateAx = plotArea.AddNewDateAx();
            ctDateAx.AddNewAxId().val = (uint)id;
            ctDateAx.AddNewAxPos();
            ctDateAx.AddNewScaling();
            ctDateAx.AddNewCrosses();
            ctDateAx.AddNewCrossAx();
            ctDateAx.AddNewTickLblPos().val = ST_TickLblPos.nextTo;
            ctDateAx.AddNewDelete();
            ctDateAx.AddNewMajorTickMark();
            ctDateAx.AddNewMinorTickMark();

            SetPosition(position);
            SetOrientation(AxisOrientation.MinMax);
            SetCrosses(AxisCrosses.AutoZero);
            SetVisible(true);
            SetMajorTickMark(AxisTickMark.Cross);
            SetMinorTickMark(AxisTickMark.None);
        }
    }
}


