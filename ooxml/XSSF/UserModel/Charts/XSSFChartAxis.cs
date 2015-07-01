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
using System;
using NPOI.OpenXmlFormats.Dml.Chart;
namespace NPOI.XSSF.UserModel.Charts
{

    /**
     * Base class for all axis types.
     *
     * @author Roman Kashitsyn
     */
    public abstract class XSSFChartAxis : IChartAxis
    {

        protected XSSFChart chart;

        private static double Min_LOG_BASE = 2.0;
        private static double Max_LOG_BASE = 1000.0;

        protected XSSFChartAxis(XSSFChart chart)
        {
            this.chart = chart;
        }
        public abstract long Id { get;}
        public abstract void CrossAxis(IChartAxis axis);

        public AxisPosition Position
        {
            get
            {
                return toAxisPosition(GetCTAxPos());
            }
            set
            {
                GetCTAxPos().val = fromAxisPosition(value);
            }
        }

        public String NumberFormat
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

        public bool IsSetLogBase
        {
            get
            {
                return GetCTScaling().IsSetLogBase();
            }
        }

        public double LogBase
        {
            get
            {
                CT_LogBase logBase = GetCTScaling().logBase;
                if (logBase != null)
                {
                    return logBase.val;
                }
                return 0.0;
            }
            set
            {
                if (value < Min_LOG_BASE ||
                Max_LOG_BASE < value)
                {
                    throw new ArgumentException("Axis log base must be between 2 and 1000 (inclusive), got: " + value);
                }
                CT_Scaling scaling = GetCTScaling();
                if (scaling.IsSetLogBase())
                {
                    scaling.logBase.val = value;
                }
                else
                {
                    scaling.AddNewLogBase().val = (value);
                }
            }
        }

        public bool IsSetMinimum
        {
            get
            {
                return GetCTScaling().IsSetMin();
            }
        }

        public double Minimum
        {
            get
            {
                CT_Scaling scaling = GetCTScaling();
                if (scaling.IsSetMin())
                {
                    return scaling.min.val;
                }
                else
                {
                    return 0.0;
                }
            }
            set
            {
                CT_Scaling scaling = GetCTScaling();
                if (scaling.IsSetMin())
                {
                    scaling.min.val = value;
                }
                else
                {
                    scaling.AddNewMin().val = value;
                }
            }
        }

        public bool IsSetMaximum
        {
            get
            {
                return GetCTScaling().IsSetMax();
            }
        }

        public double Maximum
        {
            get
            {
                CT_Scaling scaling = GetCTScaling();
                if (scaling.IsSetMax())
                {
                    return scaling.max.val;
                }
                else
                {
                    return 0.0;
                }
            }
            set
            {
                CT_Scaling scaling = GetCTScaling();
                if (scaling.IsSetMax())
                {
                    scaling.max.val = value;
                }
                else
                {
                    scaling.AddNewMax().val = value;
                }
            }
        }

        public AxisOrientation Orientation
        {
            get
            {
                return toAxisOrientation(GetCTScaling().orientation);
            }
            set
            {
                CT_Scaling scaling = GetCTScaling();
                ST_Orientation stOrientation = fromAxisOrientation(value);
                if (scaling.IsSetOrientation())
                {
                    scaling.orientation.val = stOrientation;
                }
                else
                {
                    GetCTScaling().AddNewOrientation().val = stOrientation;
                }
            }
        }

        public AxisCrosses Crosses
        {
            get
            {
                return toAxisCrosses(GetCTCrosses());
            }
            set
            {
                GetCTCrosses().val = fromAxisCrosses(value);
            }
        }
        public bool IsVisible
        {
            get
            {
                return GetDelete().val == 1;
            }
            set
            {
                GetDelete().val = value ? 1 : 0;
            }
        }

        public AxisTickMark MajorTickMark
        {
            get
            {
                return toAxisTickMark(GetMajorCTTickMark());
            }
            set
            {
                GetMajorCTTickMark().val = fromAxisTickMark(value);
            }
        }

        public AxisTickMark MinorTickMark
        {
            get
            {
                return toAxisTickMark(GetMinorCTTickMark());
            }
            set
            {
                GetMinorCTTickMark().val = fromAxisTickMark(value);
            }
        }

        protected abstract CT_AxPos GetCTAxPos();
        protected abstract CT_NumFmt GetCTNumFmt();
        protected abstract CT_Scaling GetCTScaling();
        protected abstract CT_Crosses GetCTCrosses();
        protected abstract CT_Boolean GetDelete();
        protected abstract CT_TickMark GetMajorCTTickMark();
        protected abstract CT_TickMark GetMinorCTTickMark();

        private static ST_Orientation fromAxisOrientation(AxisOrientation orientation)
        {
            switch (orientation)
            {
                case AxisOrientation.MinToMax: return ST_Orientation.minMax;
                case AxisOrientation.MaxToMin: return ST_Orientation.maxMin;
                default:
                    throw new ArgumentException();
            }
        }

        private static AxisOrientation toAxisOrientation(CT_Orientation ctOrientation)
        {
            switch (ctOrientation.val)
            {
                case ST_Orientation.minMax: return AxisOrientation.MinToMax;
                case ST_Orientation.maxMin: return AxisOrientation.MaxToMin;
                default:
                    throw new ArgumentException();
            }
        }

        private static ST_Crosses fromAxisCrosses(AxisCrosses crosses)
        {
            switch (crosses)
            {
                case AxisCrosses.AutoZero: return ST_Crosses.autoZero;
                case AxisCrosses.Min: return ST_Crosses.min;
                case AxisCrosses.Max: return ST_Crosses.max;
                default:
                    throw new ArgumentException();
            }
        }

        private static AxisCrosses toAxisCrosses(CT_Crosses ctCrosses)
        {
            switch (ctCrosses.val)
            {
                case ST_Crosses.autoZero: return AxisCrosses.AutoZero;
                case ST_Crosses.max: return AxisCrosses.Max;
                case ST_Crosses.min: return AxisCrosses.Min;
                default:
                    throw new ArgumentException();
            }
        }

        private static ST_AxPos fromAxisPosition(AxisPosition position)
        {
            switch (position)
            {
                case AxisPosition.Bottom: return ST_AxPos.b;
                case AxisPosition.Left: return ST_AxPos.l;
                case AxisPosition.Right: return ST_AxPos.r;
                case AxisPosition.Top: return ST_AxPos.t;
                default:
                    throw new ArgumentException();
            }
        }

        private static AxisPosition toAxisPosition(CT_AxPos ctAxPos)
        {
            switch (ctAxPos.val)
            {
                case ST_AxPos.b: return AxisPosition.Bottom;
                case ST_AxPos.l: return AxisPosition.Left;
                case ST_AxPos.r: return AxisPosition.Right;
                case ST_AxPos.t: return AxisPosition.Top;
                default: return AxisPosition.Bottom;
            }
        }
        private static ST_TickMark fromAxisTickMark(AxisTickMark tickMark)
        {
            switch (tickMark)
            {
                case AxisTickMark.None: return ST_TickMark.none;
                case AxisTickMark.In: return ST_TickMark.@in;
                case AxisTickMark.Out: return ST_TickMark.@out;
                case AxisTickMark.Cross: return ST_TickMark.cross;
                default:
                    throw new ArgumentException("Unknown AxisTickMark: " + tickMark);
            }
        }

        private static AxisTickMark toAxisTickMark(CT_TickMark ctTickMark)
        {
            switch (ctTickMark.val)
            {
                case ST_TickMark.none: return AxisTickMark.None;
                case ST_TickMark.@in: return AxisTickMark.In;
                case ST_TickMark.@out: return AxisTickMark.Out;
                case ST_TickMark.cross: return AxisTickMark.Cross;
                default: return AxisTickMark.Cross;
            }
        }
    }

}