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
    using NPOI.SS.UserModel;

    /**
     * High level representation for DataBar / Data-Bar Formatting 
     *  component of Conditional Formatting Settings
     */
    public class HSSFDataBarFormatting : IDataBarFormatting
    {
        private HSSFSheet sheet;
        private CFRule12Record cfRule12Record;
        private DataBarFormatting databarFormatting;

        protected internal HSSFDataBarFormatting(CFRule12Record cfRule12Record, HSSFSheet sheet)
        {
            this.sheet = sheet;
            this.cfRule12Record = cfRule12Record;
            this.databarFormatting = this.cfRule12Record.DataBarFormatting;
        }

        public bool IsLeftToRight
        {
            get
            {
                return !databarFormatting.IsReversed;
            }
            set
            {
                databarFormatting.IsReversed = value;
            }
        }

        public int WidthMin
        {
            get
            {
                return databarFormatting.PercentMin;
            }
            set
            {
                databarFormatting.PercentMin = (byte)value;
            }
        }

        public int WidthMax
        {
            get
            {
                return databarFormatting.PercentMax;
            }
            set
            {
                databarFormatting.PercentMax = (byte)value;
            }
        }

        public IColor Color
        {
            get
            {
                return new HSSFExtendedColor(databarFormatting.Color);
            }
            set
            {
                HSSFExtendedColor hcolor = (HSSFExtendedColor)value;
                databarFormatting.Color = (/*setter*/hcolor.ExtendedColor);
            }
        }

        public IConditionalFormattingThreshold MinThreshold
        {
            get
            {
                return new HSSFConditionalFormattingThreshold(databarFormatting.ThresholdMin, sheet);
            }
        }
        public IConditionalFormattingThreshold MaxThreshold
        {
            get
            {
                return new HSSFConditionalFormattingThreshold(databarFormatting.ThresholdMax, sheet);
            }
        }

        public bool IsIconOnly
        {
            get
            {
                return databarFormatting.IsIconOnly;
            }
            set
            {
                databarFormatting.IsIconOnly = value;
            }
        }

        public HSSFConditionalFormattingThreshold CreateThreshold()
        {
            return new HSSFConditionalFormattingThreshold(new DataBarThreshold(), sheet);
        }
    }

}