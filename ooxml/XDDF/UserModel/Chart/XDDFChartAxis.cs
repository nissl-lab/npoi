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
    /// Base class for all axis types.
    /// </summary>
    public abstract class XDDFChartAxis : IHasShapeProperties
    {
        protected abstract CT_UnsignedInt GetCTAxId();

        protected abstract CT_AxPos GetCTAxPos();

        protected abstract CT_NumFmt GetCTNumFmt();

        protected abstract CT_Scaling GetCTScaling();

        protected abstract CT_Crosses GetCTCrosses();

        protected abstract CT_Boolean GetDelete();

        protected abstract CT_TickMark GetMajorCTTickMark();

        protected abstract CT_TickMark GetMinorCTTickMark();

        public abstract XDDFShapeProperties GetOrAddMajorGridProperties();

        public abstract XDDFShapeProperties GetOrAddMinorGridProperties();

        /// <summary>
        /// </summary>
        /// <returns>true if minor unit value is defined, false otherwise</returns>
        public abstract bool IsSetMinorUnit();

        /// <summary>
        /// </summary>
        /// <param name="minor">
        /// axis minor unit
        /// </param>
        public abstract void SetMinorUnit(double minor);

        /// <summary>
        /// </summary>
        /// <returns>axis minor unit or NaN if not Set</returns>
        public abstract double GetMinorUnit();

        /// <summary>
        /// </summary>
        /// <returns>true if major unit value is defined, false otherwise</returns>
        public abstract bool IsSetMajorUnit();

        /// <summary>
        /// </summary>
        /// <param name="major">
        /// axis major unit
        /// </param>
        public abstract void SetMajorUnit(double major);

        /// <summary>
        /// </summary>
        /// <returns>axis major unit or NaN if not Set</returns>
        public abstract double GetMajorUnit();

        /// <summary>
        /// </summary>
        /// <returns>axis id</returns>
        public long GetId()
        {
            return GetCTAxId().val;
        }

        /// <summary>
        /// </summary>
        /// <returns>axis position</returns>
        public AxisPosition GetPosition()
        {
            return AxisPositionExtensions.ValueOf(GetCTAxPos().val);
        }

        /// <summary>
        /// </summary>
        /// <param name="position">
        /// new axis position
        /// </param>
        public void SetPosition(AxisPosition position)
        {
            GetCTAxPos().val = position.ToST_AxPos();
        }

        /// <summary>
        /// Use this to check before retrieving a number format, as calling
        /// <see cref="getNumberFormat()" /> may create a default one if none exists.
        /// </summary>
        /// <returns>true if a number format element is defined, false if not</returns>
        public abstract bool HasNumberFormat();

        /// <summary>
        /// </summary>
        /// <param name="format">
        /// axis number format
        /// </param>
        public void SetNumberFormat(string format)
        {
            GetCTNumFmt().formatCode = format;
            GetCTNumFmt().sourceLinked = true;
        }

        /// <summary>
        /// </summary>
        /// <returns>axis number format</returns>
        public string GetNumberFormat()
        {
            return GetCTNumFmt().formatCode;
        }

        /// <summary>
        /// </summary>
        /// <returns>true if log base is defined, false otherwise</returns>
        public bool IsSetLogBase()
        {
            return GetCTScaling().IsSetLogBase();
        }

        private static  double MIN_LOG_BASE = 2.0;
        private static  double MAX_LOG_BASE = 1000.0;

        /// <summary>
        /// </summary>
        /// <param name="logBase">
        /// a number between 2 and 1000 (inclusive)
        /// </param>
        /// <exception cref="ArgumentException">
        /// if log base not within allowed range
        /// </exception>
        public void SetLogBase(double logBase)
        {
            if(logBase < MIN_LOG_BASE || MAX_LOG_BASE < logBase)
            {
                throw new ArgumentException("Axis log base must be between 2 and 1000 (inclusive), got: " + logBase);
            }
            CT_Scaling scaling = GetCTScaling();
            if(scaling.IsSetLogBase())
            {
                scaling.logBase.val = logBase;
            }
            else
            {
                scaling.AddNewLogBase().val = logBase;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>axis log base or NaN if not Set</returns>
        public double GetLogBase()
        {
            CT_Scaling scaling = GetCTScaling();
            if(scaling.IsSetLogBase())
            {
                return scaling.logBase.val;
            }
            return Double.NaN;
        }

        /// <summary>
        /// </summary>
        /// <returns>true if minimum value is defined, false otherwise</returns>
        public bool IsSetMinimum()
        {
            return GetCTScaling().IsSetMin();
        }

        /// <summary>
        /// </summary>
        /// <param name="min">
        /// axis minimum
        /// </param>
        public void SetMinimum(double min)
        {
            CT_Scaling scaling = GetCTScaling();
            if(Double.IsNaN(min))
            {
                if(scaling.IsSetMin())
                {
                    scaling.UnsetMin();
                }
            }
            else
            {
                if(scaling.IsSetMin())
                {
                    scaling.min.val = min;
                }
                else
                {
                    scaling.AddNewMin().val = min;
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>axis minimum or NaN if not Set</returns>
        public double GetMinimum()
        {
            CT_Scaling scaling = GetCTScaling();
            if(scaling.IsSetMin())
            {
                return scaling.min.val;
            }
            else
            {
                return Double.NaN;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if maximum value is defined, false otherwise</returns>
        public bool IsSetMaximum()
        {
            return GetCTScaling().IsSetMax();
        }

        /// <summary>
        /// </summary>
        /// <param name="max">
        /// axis maximum
        /// </param>
        public void SetMaximum(double max)
        {
            CT_Scaling scaling = GetCTScaling();
            if(Double.IsNaN(max))
            {
                if(scaling.IsSetMax())
                {
                    scaling.UnsetMax();
                }
            }
            else
            {
                if(scaling.IsSetMax())
                {
                    scaling.max.val = max;
                }
                else
                {
                    scaling.AddNewMax().val = max;
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>axis maximum or NaN if not Set</returns>
        public double GetMaximum()
        {
            CT_Scaling scaling = GetCTScaling();
            if(scaling.IsSetMax())
            {
                return scaling.max.val;
            }
            else
            {
                return Double.NaN;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>axis orientation</returns>
        public AxisOrientation GetOrientation()
        {
            return AxisOrientationExtensions.ValueOf(GetCTScaling().orientation.val);
        }

        /// <summary>
        /// </summary>
        /// <param name="orientation">
        /// axis orientation
        /// </param>
        public void SetOrientation(AxisOrientation orientation)
        {
            CT_Scaling scaling = GetCTScaling();
            if(scaling.IsSetOrientation())
            {
                scaling.orientation.val = orientation.ToST_Orientation();
            }
            else
            {
                scaling.AddNewOrientation().val = orientation.ToST_Orientation();
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>axis cross type</returns>
        public AxisCrosses GetCrosses()
        {
            return AxisCrossesExtensions.ValueOf(GetCTCrosses().val);
        }

        /// <summary>
        /// </summary>
        /// <param name="crosses">
        /// axis cross type
        /// </param>
        public void SetCrosses(AxisCrosses crosses)
        {
            GetCTCrosses().val = crosses.ToST_Crosses();
        }

        /// <summary>
        /// Declare this axis cross another axis.
        /// </summary>
        /// <param name="axis">
        /// that this axis should cross
        /// </param>
        public abstract void crossAxis(XDDFChartAxis axis);

        /// <summary>
        /// </summary>
        /// <returns>visibility of the axis.</returns>
        public bool IsVisible()
        {
            return !(GetDelete().val > 0);
        }

        /// <summary>
        /// </summary>
        /// <param name="value">
        /// visibility of the axis.
        /// </param>
        public void SetVisible(bool value)
        {
            GetDelete().val = value ? 0 : 1;
        }

        /// <summary>
        /// </summary>
        /// <returns>major tick mark.</returns>
        public AxisTickMark GetMajorTickMark()
        {
            return AxisTickMarkExtensions.ValueOf(GetMajorCTTickMark().val);
        }

        /// <summary>
        /// </summary>
        /// <param name="tickMark">
        /// major tick mark type.
        /// </param>
        public void SetMajorTickMark(AxisTickMark tickMark)
        {
            GetMajorCTTickMark().val = tickMark.ToST_TickMark();
        }

        /// <summary>
        /// </summary>
        /// <returns>minor tick mark.</returns>
        public AxisTickMark GetMinorTickMark()
        {
            return AxisTickMarkExtensions.ValueOf(GetMinorCTTickMark().val);
        }

        /// <summary>
        /// </summary>
        /// <param name="tickMark">
        /// minor tick mark type.
        /// </param>
        public void SetMinorTickMark(AxisTickMark tickMark)
        {
            GetMinorCTTickMark().val =tickMark.ToST_TickMark();
        }

        protected OpenXmlFormats.Dml.Chart.CT_ShapeProperties GetOrAddLinesProperties(CT_ChartLines gridlines)
        {
            OpenXmlFormats.Dml.Chart.CT_ShapeProperties properties;
            if(gridlines.IsSetSpPr())
            {
                properties = gridlines.spPr;
            }
            else
            {
                properties = gridlines.AddNewSpPr();
            }
            return properties;
        }

        protected long GetNextAxId(CT_PlotArea plotArea)
        {
            long totalAxisCount = plotArea.SizeOfValAxArray() + plotArea.SizeOfCatAxArray() + plotArea.SizeOfDateAxArray()
                + plotArea.SizeOfSerAxArray();
            return totalAxisCount;
        }

        public abstract XDDFShapeProperties GetOrAddShapeProperties();
    }
}


