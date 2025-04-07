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
using System.Collections.Generic;
using System.Linq;

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.OpenXmlFormats.Dml.Chart;
    using NPOI.Util.Optional;
    using NPOI.XDDF.UserModel.Text;


    /// <summary>
    /// Represents a DrawingML chart legend
    /// </summary>
    public sealed class XDDFChartLegend : ITextContainer
    {

        /// <summary>
        /// Underlying CTLegend bean
        /// </summary>
        private CT_Legend legend;

        /// <summary>
        /// Create a new DrawingML chart legend
        /// </summary>
        public XDDFChartLegend(CT_Chart ctChart)
        {
            this.legend = (ctChart.IsSetLegend()) ? ctChart.legend : ctChart.AddNewLegend();

            SetDefaults();
        }

        /// <summary>
        /// Set sensible default styling.
        /// </summary>
        private void SetDefaults()
        {
            if(!legend.IsSetOverlay())
            {
                legend.AddNewOverlay();
            }
            legend.overlay.val = 0;
        }

        /// <summary>
        /// Return the underlying CTLegend bean.
        /// </summary>
        /// <returns>the underlying CTLegend bean</returns>
        protected CT_Legend GetXmlobject()
        {
            return legend;
        }

        // will later replace with XDDFShapeProperties
        public CT_ShapeProperties ShapeProperties
        {
            get
            {
                if(legend.IsSetSpPr())
                {
                    return legend.spPr;
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
                    if(legend.IsSetSpPr())
                    {
                        legend.UnsetSpPr();
                    }
                }
                else
                {
                    legend.spPr = value;
                }
            }
        }



        public XDDFTextBody TextBody
        {
            get
            {
                if(legend.IsSetTxPr())
                {
                    return new XDDFTextBody(this, legend.txPr);
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
                    if(legend.IsSetTxPr())
                    {
                        legend.UnsetTxPr();
                    }
                }
                else
                {
                    legend.txPr = value.GetXmlObject();
                }
            }
        }

        public XDDFLegendEntry AddEntry()
        {
            return new XDDFLegendEntry(legend.AddNewLegendEntry());
        }

        public XDDFLegendEntry GetEntry(int index)
        {
            return new XDDFLegendEntry(legend.GetLegendEntryArray(index));
        }

        public List<XDDFLegendEntry> GetEntries()
        {
            return [.. legend.legendEntry.Select(x => new XDDFLegendEntry(x))];
        }

        public XDDFChartExtensionList ExtensionList
        {
            get
            {
                if(legend.IsSetExtLst())
                {
                    return new XDDFChartExtensionList(legend.extLst);
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
                    if(legend.IsSetExtLst())
                    {
                        legend.UnsetExtLst();
                    }
                }
                else
                {
                    legend.extLst = value.GetXmlObject();
                }
            }
        }

        public XDDFLayout Layout
        {
            get
            {
                if(legend.IsSetLayout())
                {
                    return new XDDFLayout(legend.layout);
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
                    if(legend.IsSetLayout())
                    {
                        legend.UnsetLayout();
                    }
                }
                else
                {
                    legend.layout = value.GetXmlObject();
                }
            }
        }

        public void SetPosition(LegendPosition position)
        {

        }

        /*
         * According to ECMA-376 default position is RIGHT.
         */
        public LegendPosition Position
        {
            get
            {
                if(legend.IsSetLegendPos())
                {
                    return LegendPositionExtensions.ValueOf(legend.legendPos.val);
                }
                else
                {
                    return LegendPosition.Right;
                }
            }
            set
            {
                if(!legend.IsSetLegendPos())
                {
                    legend.AddNewLegendPos();
                }
                legend.legendPos.val = value.ToST_LegendPos();
            }
        }

        public XDDFManualLayout GetOrAddManualLayout()
        {
            if(!legend.IsSetLayout())
            {
                legend.AddNewLayout();
            }
            return new XDDFManualLayout(legend.layout);
        }

        public bool IsOverlay
        {
            get => legend.overlay.val == 1;
            set => legend.overlay.val = value ? 1 : 0;
        }



        public Option<R> FindDefinedParagraphProperty<R>(
                    Func<CT_TextParagraphProperties, Boolean> isSet,
                    Func<CT_TextParagraphProperties, R> Getter) where R : class
        {
            return Option<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public Option<R> FindDefinedRunProperty<R>(
                Func<CT_TextCharacterProperties, Boolean> isSet,
                Func<CT_TextCharacterProperties, R> Getter) where R : class
        {
            return Option<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public ValueOption<R> FindDefinedParagraphValueProperty<R>(
                Func<CT_TextParagraphProperties, Boolean> isSet,
                Func<CT_TextParagraphProperties, R> Getter) where R : struct
        {
            return ValueOption<R>.None(); // legend entry has no (indirect) paragraph properties
        }

        public ValueOption<R> FindDefinedRunValueProperty<R>(
                Func<CT_TextCharacterProperties, Boolean> isSet,
                Func<CT_TextCharacterProperties, R> Getter) where R : struct
        {
            return ValueOption<R>.None(); // legend entry has no (indirect) paragraph properties
        }
    }
}


