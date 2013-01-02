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
     * Represents a SpreadsheetML chart legend
     * @author Roman Kashitsyn
     */
    public class XSSFChartLegend : IChartLegend
    {

        /**
         * Underlaying CTLagend bean
         */
        private CT_Legend legend;

        /**
         * Create a new SpreadsheetML chart legend
         */
        public XSSFChartLegend(XSSFChart chart)
        {
            CT_Chart ctChart = chart.GetCTChart();
            this.legend = (ctChart.IsSetLegend()) ?
                ctChart.legend :
                ctChart.AddNewLegend();
        }

        /**
         * Return the underlying CTLegend bean.
         *
         * @return the underlying CTLegend bean
         */

        internal CT_Legend GetCTLegend()
        {
            return legend;
        }
        
        /*
         * According to ECMA-376 default position is RIGHT.
         */
        public LegendPosition Position
        {
            get
            {
                if (legend.IsSetLegendPos())
                {
                    return ToLegendPosition(legend.legendPos);
                }
                else
                {
                    return LegendPosition.RIGHT;
                }
            }
            set 
            {
                if (!legend.IsSetLegendPos())
                {
                    legend.AddNewLegendPos();
                }
                legend.legendPos.val = FromLegendPosition(value);
                legend.legendPosSpecified = true;
            }
        }

        public IManualLayout GetManualLayout()
        {
            if (!legend.IsSetLayout())
            {
                legend.AddNewLayout();
            }
            return new XSSFManualLayout(legend.layout);
        }

        private ST_LegendPos FromLegendPosition(LegendPosition position)
        {
            switch (position)
            {
                case LegendPosition.BOTTOM: return ST_LegendPos.b;
                case LegendPosition.LEFT: return ST_LegendPos.l;
                case LegendPosition.RIGHT: return ST_LegendPos.r;
                case LegendPosition.TOP: return ST_LegendPos.t;
                case LegendPosition.TOP_RIGHT: return ST_LegendPos.tr;
                default:
                    throw new ArgumentException();
            }
        }

        private LegendPosition ToLegendPosition(CT_LegendPos ctLegendPos)
        {
            switch (ctLegendPos.val)
            {
                case ST_LegendPos.b: return LegendPosition.BOTTOM;
                case ST_LegendPos.l: return LegendPosition.LEFT;
                case ST_LegendPos.r: return LegendPosition.RIGHT;
                case ST_LegendPos.t: return LegendPosition.TOP;
                case ST_LegendPos.tr: return LegendPosition.TOP_RIGHT;
                default:
                    throw new ArgumentException();
            }
        }
    }
}

