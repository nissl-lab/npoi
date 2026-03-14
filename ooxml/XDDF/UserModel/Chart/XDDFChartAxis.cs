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
        public long Id
        {
            get
            {
                return GetCTAxId().val;
            }
        }

        /// <summary>
        /// axis position
        /// </summary>
        public AxisPosition Position
        {
            get
            {
                return AxisPositionExtensions.ValueOf(GetCTAxPos().val);
            }
            set 
            {
                GetCTAxPos().val = value.ToST_AxPos();
            }
        }

        /// <summary>
        /// Use this to check before retrieving a number format, as calling
        /// <see cref="getNumberFormat()" /> may create a default one if none exists.
        /// </summary>
        /// <returns>true if a number format element is defined, false if not</returns>
        public abstract bool HasNumberFormat();

        /// <summary>
        /// axis number format
        /// </summary>
        public string NumberFormat
        {
            get
            {
                return GetCTNumFmt().formatCode;
            }
            set 
            {
                GetCTNumFmt().formatCode = value;
                GetCTNumFmt().sourceLinked = true;
            }
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
        /// axis log base or Double.NaN if not set,  a number between 2 and 1000 (inclusive)
        /// </summary>
        public double LogBase
        {
            get
            {
                CT_Scaling scaling = GetCTScaling();
                if(scaling.IsSetLogBase())
                {
                    return scaling.logBase.val;
                }
                return Double.NaN;
            }
            set 
            {
                if(value < MIN_LOG_BASE || MAX_LOG_BASE < value)
                {
                    throw new ArgumentException("Axis log base must be between 2 and 1000 (inclusive), got: " + value);
                }
                CT_Scaling scaling = GetCTScaling();
                if(scaling.IsSetLogBase())
                {
                    scaling.logBase.val = value;
                }
                else
                {
                    scaling.AddNewLogBase().val = value;
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if minimum value is defined, false otherwise</returns>
        public bool IsSetMinimum()
        {
            return GetCTScaling().IsSetMin();
        }

        /// <summary>
        /// axis minimum or NaN if not set
        /// </summary>
        public double Minimum
        {
            get
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
            set 
            {
                CT_Scaling scaling = GetCTScaling();
                if(Double.IsNaN(value))
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
                        scaling.min.val = value;
                    }
                    else
                    {
                        scaling.AddNewMin().val = value;
                    }
                }
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
        /// axis maximum or double.NaN if not set
        /// </summary>
        public double Maximum
        {
            get
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
            set 
            {
                CT_Scaling scaling = GetCTScaling();
                if(Double.IsNaN(value))
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
                        scaling.max.val = value;
                    }
                    else
                    {
                        scaling.AddNewMax().val = value;
                    }
                }
            }
        }

        /// <summary>
        /// axis orientation
        /// </summary>
        public AxisOrientation Orientation
        {
            get
            {
                return AxisOrientationExtensions.ValueOf(GetCTScaling().orientation.val);
            }
            set
            {
                CT_Scaling scaling = GetCTScaling();
                if(scaling.IsSetOrientation())
                {
                    scaling.orientation.val = value.ToST_Orientation();
                }
                else
                {
                    scaling.AddNewOrientation().val = value.ToST_Orientation();
                }
            }
        }

        /// <summary>
        /// axis cross type
        /// </summary>
        public AxisCrosses Crosses
        {
            get
            {
                return AxisCrossesExtensions.ValueOf(GetCTCrosses().val);
            }
            set 
            {
                GetCTCrosses().val = value.ToST_Crosses();
            }
        }

        /// <summary>
        /// Declare this axis cross another axis.
        /// </summary>
        /// <param name="axis">
        /// that this axis should cross
        /// </param>
        public abstract void CrossAxis(XDDFChartAxis axis);


        /// <summary>
        /// visibility of the axis.
        /// </summary>
        public bool IsVisible
        {
            get
            {
                return !(GetDelete().val > 0);
            }
            set
            {
                GetDelete().val = value ? 0 : 1;
            }
        }

        /// <summary>
        /// major tick mark.
        /// </summary>
        public AxisTickMark MajorTickMark
        {
            get
            {
                return AxisTickMarkExtensions.ValueOf(GetMajorCTTickMark().val);
            }
            set 
            {
                GetMajorCTTickMark().val = value.ToST_TickMark();
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>minor tick mark.</returns>
        public AxisTickMark MinorTickMark
        {
            get
            {
                return AxisTickMarkExtensions.ValueOf(GetMinorCTTickMark().val);
            }
            set 
            {
                GetMinorCTTickMark().val =value.ToST_TickMark();
            }
        }

        protected CT_ShapeProperties GetOrAddLinesProperties(CT_ChartLines gridlines)
        {
            CT_ShapeProperties properties;
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


