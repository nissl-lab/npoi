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

namespace NPOI.HSSF.UserModel
{
    using System;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.Record.Common;
    using NPOI.SS.UserModel;
    using ExtendedColorR = NPOI.HSSF.Record.Common.ExtendedColor;
    /**
     * High level representation for Color Scale / Color Gradient 
     *  Formatting component of Conditional Formatting Settings
     */
    public class HSSFColorScaleFormatting : IColorScaleFormatting
    {
        private HSSFSheet sheet;
        private CFRule12Record cfRule12Record;
        private ColorGradientFormatting colorFormatting;

        protected internal HSSFColorScaleFormatting(CFRule12Record cfRule12Record, HSSFSheet sheet)
        {
            this.sheet = sheet;
            this.cfRule12Record = cfRule12Record;
            this.colorFormatting = this.cfRule12Record.ColorGradientFormatting;
        }

        public int NumControlPoints
        {
            get
            {
                return colorFormatting.NumControlPoints;
            }
            set
            {
                colorFormatting.NumControlPoints = (value);
            }
        }

        public IColor[] Colors
        {
            get
            {
                ExtendedColorR[] colors = colorFormatting.Colors;
                HSSFExtendedColor[] hcolors = new HSSFExtendedColor[colors.Length];
                for (int i = 0; i < colors.Length; i++)
                {
                    hcolors[i] = new HSSFExtendedColor(colors[i]);
                }
                return hcolors;
            }
            set
            {
                ExtendedColorR[] cr = new ExtendedColorR[value.Length];
                for (int i = 0; i < value.Length; i++)
                {
                    cr[i] = ((HSSFExtendedColor)value[i]).ExtendedColor;
                }
                colorFormatting.Colors = (/*setter*/cr);
            }
        }

        public IConditionalFormattingThreshold[] Thresholds
        {
            get
            {
                Threshold[] t = colorFormatting.Thresholds;
                HSSFConditionalFormattingThreshold[] ht = new HSSFConditionalFormattingThreshold[t.Length];
                for (int i = 0; i < t.Length; i++)
                {
                    ht[i] = new HSSFConditionalFormattingThreshold(t[i], sheet);
                }
                return ht;
            }
            set
            {
                ColorGradientThreshold[] t = new ColorGradientThreshold[value.Length];
                for (int i = 0; i < t.Length; i++)
                {
                    HSSFConditionalFormattingThreshold hssfT = (HSSFConditionalFormattingThreshold)value[i];
                    t[i] = (ColorGradientThreshold)hssfT.Threshold;
                }
                colorFormatting.Thresholds = (/*setter*/t);
            }
        }

        public IConditionalFormattingThreshold CreateThreshold()
        {
            return new HSSFConditionalFormattingThreshold(new ColorGradientThreshold(), sheet);
        }
    }

}