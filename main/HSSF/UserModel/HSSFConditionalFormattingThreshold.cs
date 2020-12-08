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

    using NPOI.HSSF.Record.CF;
    using NPOI.SS.UserModel;

    /**
     * High level representation for Icon / Multi-State / Databar /
     *  Colour Scale change thresholds
     */
    public class HSSFConditionalFormattingThreshold : IConditionalFormattingThreshold
    {
        private Threshold threshold;
        private HSSFSheet sheet;
        private HSSFWorkbook workbook;

        protected internal HSSFConditionalFormattingThreshold(Threshold threshold, HSSFSheet sheet)
        {
            this.threshold = threshold;
            this.sheet = sheet;
            this.workbook = sheet.Workbook as HSSFWorkbook;
        }
        protected internal Threshold Threshold
        {
            get { return threshold; }
        }

        public RangeType RangeType
        {
            get
            {
                return RangeType.ById(threshold.Type);
            }
            set
            {
                threshold.Type = (byte)value.id;
            }
        }


        public String Formula
        {
            get { return HSSFConditionalFormattingRule.ToFormulaString(threshold.ParsedExpression, workbook); }
            set { threshold.ParsedExpression = NPOI.HSSF.Record.CFRuleBase.ParseFormula(value, sheet); }
        }

        public double? Value
        {
            get { return threshold.Value; }
            set { threshold.Value = (value); }
        }
    }

}