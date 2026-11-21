/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.	See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.	 You may obtain a copy of the License at

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
    using NPOI.OpenXmlFormats.Dml.Chart;
    /// <summary>
    /// Represents a DrawingML manual layout.
    /// </summary>
    public sealed class XDDFManualLayout
    {

        /// <summary>
        /// Underlaying CTManualLayout bean.
        /// </summary>
        private CT_ManualLayout layout;

        private static LayoutMode defaultLayoutMode = LayoutMode.Edge;
        private static LayoutTarget defaultLayoutTarget = LayoutTarget.Inner;

        /// <summary>
        /// Create a new DrawingML manual layout.
        /// </summary>
        /// <param name="ctLayout">
        /// a DrawingML layout that should be used as base.
        /// </param>
        public XDDFManualLayout(CT_Layout ctLayout)
        {
            InitializeLayout(ctLayout);
        }

        /// <summary>
        /// Create a new DrawingML manual layout for chart.
        /// </summary>
        /// <param name="ctPlotArea">
        /// a chart's plot area to create layout for.
        /// </param>
        public XDDFManualLayout(CT_PlotArea ctPlotArea)
        {
            CT_Layout ctLayout = ctPlotArea.IsSetLayout() ? ctPlotArea.layout : ctPlotArea.AddNewLayout();

            InitializeLayout(ctLayout);
        }

        /// <summary>
        /// Return the underlying CTManualLayout bean.
        /// </summary>
        /// <returns>the underlying CTManualLayout bean.</returns>
        internal CT_ManualLayout GetXmlObject()
        {
            return layout;
        }

        public XDDFChartExtensionList ExtensionList
        {
            get
            {
                if(layout.IsSetExtLst())
                {
                    return new XDDFChartExtensionList(layout.extLst);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    if(layout.IsSetExtLst())
                    {
                        layout.UnsetExtLst();
                    }
                }
                else
                {
                    layout.extLst = value.GetXmlObject();
                }
            }
        }

        public double WidthRatio
        {
            get
            {
                if(!layout.IsSetW())
                {
                    return 0.0;
                }
                return layout.w.val;
            }
            set
            {
                if(!layout.IsSetW())
                {
                    layout.AddNewW();
                }
                layout.w.val = value;
            }
        }

        public double HeightRatio
        {
            get
            {
                if(!layout.IsSetH())
                {
                    return 0.0;
                }
                return layout.h.val;
            }
            set
            {
                if(!layout.IsSetH())
                {
                    layout.AddNewH();
                }
                layout.h.val = value;
            }
        }

        public LayoutTarget Target
        {
            get
            {
                if(!layout.IsSetLayoutTarget())
                {
                    return defaultLayoutTarget;
                }
                return LayoutTargetExtensions.ValueOf(layout.layoutTarget.val);
            }
            set
            {
                if(!layout.IsSetLayoutTarget())
                {
                    layout.AddNewLayoutTarget();
                }
                layout.layoutTarget.val = value.ToST_LayoutTarget();
            }
        }

        public LayoutMode XMode
        {
            get
            {
                if(!layout.IsSetXMode())
                {
                    return defaultLayoutMode;
                }
                return LayoutModeExtensions.ValueOf(layout.xMode.val);
            }
            set
            {
                if(!layout.IsSetXMode())
                {
                    layout.AddNewXMode();
                }
                layout.xMode.val = value.ToST_LayoutMode();
            }
        }

        public LayoutMode YMode
        {
            get
            {
                if(!layout.IsSetYMode())
                {
                    return defaultLayoutMode;
                }
                return LayoutModeExtensions.ValueOf(layout.yMode.val);
            }
            set
            {
                if(!layout.IsSetYMode())
                {
                    layout.AddNewYMode();
                }
                layout.yMode.val = value.ToST_LayoutMode();
            }
        }


        public double X
        {
            get
            {
                if(!layout.IsSetX())
                {
                    return 0.0;
                }
                return layout.x.val;
            }
            set
            {
                if(!layout.IsSetX())
                {
                    layout.AddNewX();
                }
                layout.x.val = value;
            }
        }

        public double Y
        {
            get
            {
                if(!layout.IsSetY())
                {
                    return 0.0;
                }
                return layout.y.val;
            }
            set
            {
                if(!layout.IsSetY())
                {
                    layout.AddNewY();
                }
                layout.y.val = value;
            }
        }

        public LayoutMode WidthMode
        {
            get
            {
                if(!layout.IsSetWMode())
                {
                    return defaultLayoutMode;
                }
                return LayoutModeExtensions.ValueOf(layout.wMode.val);
            }
            set
            {
                if(!layout.IsSetWMode())
                {
                    layout.AddNewWMode();
                }
                layout.wMode.val = value.ToST_LayoutMode();
            }
        }

        public LayoutMode HeightMode
        {
            get
            {
                if(!layout.IsSetHMode())
                {
                    return defaultLayoutMode;
                }
                return LayoutModeExtensions.ValueOf(layout.hMode.val);
            }
            set
            {
                if(!layout.IsSetHMode())
                {
                    layout.AddNewHMode();
                }
                layout.hMode.val = value.ToST_LayoutMode();
            }
        }

        private void InitializeLayout(CT_Layout ctLayout)
        {
            if(ctLayout.IsSetManualLayout())
            {
                this.layout = ctLayout.manualLayout;
            }
            else
            {
                this.layout = ctLayout.AddNewManualLayout();
            }
        }
    }
}
