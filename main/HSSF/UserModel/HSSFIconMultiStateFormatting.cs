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
     * High level representation for Icon / Multi-State Formatting 
     *  component of Conditional Formatting Settings
     */
    public class HSSFIconMultiStateFormatting : IIconMultiStateFormatting
    {
        private HSSFSheet sheet;
        private CFRule12Record cfRule12Record;
        private IconMultiStateFormatting iconFormatting;

        protected internal HSSFIconMultiStateFormatting(CFRule12Record cfRule12Record, HSSFSheet sheet)
        {
            this.sheet = sheet;
            this.cfRule12Record = cfRule12Record;
            this.iconFormatting = this.cfRule12Record.MultiStateFormatting;
        }

        public IconSet IconSet
        {
            get { return iconFormatting.IconSet; }
            set { iconFormatting.IconSet = (value); }
        }

        public bool IsIconOnly
        {
            get { return iconFormatting.IsIconOnly; }
            set { iconFormatting.IsIconOnly = (value); }
        }

        public bool IsReversed
        {
            get { return iconFormatting.IsReversed; }
            set { iconFormatting.IsReversed = (value); }
        }


        public IConditionalFormattingThreshold[] Thresholds
        {
            get
            {
                Threshold[] t = iconFormatting.Thresholds;
                HSSFConditionalFormattingThreshold[] ht = new HSSFConditionalFormattingThreshold[t.Length];
                for (int i = 0; i < t.Length; i++)
                {
                    ht[i] = new HSSFConditionalFormattingThreshold(t[i], sheet);
                }
                return ht;
            }
            set
            {
                Threshold[] t = new Threshold[value.Length];
                for (int i = 0; i < t.Length; i++)
                {
                    t[i] = ((HSSFConditionalFormattingThreshold)value[i]).Threshold;
                }
                iconFormatting.Thresholds = t;
            }
        }


        public IConditionalFormattingThreshold CreateThreshold()
        {
            return new HSSFConditionalFormattingThreshold(new IconMultiStateThreshold(), sheet);
        }
    }

}