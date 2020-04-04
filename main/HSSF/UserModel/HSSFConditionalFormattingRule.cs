/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{

    using System;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.CF;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;

    /**
     * 
     * High level representation of Conditional Formatting Rule.
     * It allows to specify formula based conditions for the Conditional Formatting
     * and the formatting Settings such as font, border and pattern.
     * 
     * @author Dmitriy Kumshayev
     */

    public class HSSFConditionalFormattingRule : IConditionalFormattingRule
    {
        private const byte CELL_COMPARISON = CFRuleRecord.CONDITION_TYPE_CELL_VALUE_IS;

        private CFRuleBase cfRuleRecord;
        private HSSFWorkbook workbook;
        private HSSFSheet sheet;

        public HSSFConditionalFormattingRule(HSSFSheet pSheet, CFRuleBase pRuleRecord)
        {
            if (pSheet == null)
            {
                throw new ArgumentException("pSheet must not be null");
            }
            if (pRuleRecord == null)
            {
                throw new ArgumentException("pRuleRecord must not be null");
            }
            sheet = pSheet;
            workbook = pSheet.Workbook as HSSFWorkbook;
            cfRuleRecord = pRuleRecord;
        }

        public CFRuleBase CfRuleRecord
        {
            get { return cfRuleRecord; }
        }
        private CFRule12Record GetCFRule12Record(bool create)
        {
            if (cfRuleRecord is CFRule12Record)
            {
                // Good
            }
            else
            {
                if (create) throw new ArgumentException("Can't convert a CF into a CF12 record");
                return null;
            }
            return (CFRule12Record)cfRuleRecord;
        }
        private HSSFFontFormatting GetFontFormatting(bool Create)
        {
            FontFormatting fontFormatting = cfRuleRecord.FontFormatting;
            if (fontFormatting != null)
            {
                cfRuleRecord.FontFormatting=(fontFormatting);
                return new HSSFFontFormatting(cfRuleRecord, workbook);
            }
            else if (Create)
            {
                fontFormatting = new FontFormatting();
                cfRuleRecord.FontFormatting=(fontFormatting);
                return new HSSFFontFormatting(cfRuleRecord, workbook);
            }
            else
            {
                return null;
            }
        }

        /**
         * @return - font formatting object  if defined,  <c>null</c> otherwise
         */
        public IFontFormatting FontFormatting
        {
            get { return GetFontFormatting(false); }
        }
        /**
         * Create a new font formatting structure if it does not exist, 
         * otherwise just return existing object.
         * @return - font formatting object, never returns <c>null</c>. 
         */
        public IFontFormatting CreateFontFormatting()
        {
            return GetFontFormatting(true);
        }

        private HSSFBorderFormatting GetBorderFormatting(bool Create)
        {
            BorderFormatting borderFormatting = cfRuleRecord.BorderFormatting;
            if (borderFormatting != null)
            {
                cfRuleRecord.BorderFormatting=(borderFormatting);
                return new HSSFBorderFormatting(cfRuleRecord, workbook);
            }
            else if (Create)
            {
                borderFormatting = new BorderFormatting();
                cfRuleRecord.BorderFormatting=(borderFormatting);
                return new HSSFBorderFormatting(cfRuleRecord, workbook);
            }
            else
            {
                return null;
            }
        }
        /**
         * @return - border formatting object  if defined,  <c>null</c> otherwise
         */
        public IBorderFormatting BorderFormatting
        {
            get { return GetBorderFormatting(false); }
        }
        /**
         * Create a new border formatting structure if it does not exist, 
         * otherwise just return existing object.
         * @return - border formatting object, never returns <c>null</c>. 
         */
        public IBorderFormatting CreateBorderFormatting()
        {
            return GetBorderFormatting(true);
        }

        private HSSFPatternFormatting GetPatternFormatting(bool Create)
        {
            PatternFormatting patternFormatting = cfRuleRecord.PatternFormatting;
            if (patternFormatting != null)
            {
                cfRuleRecord.PatternFormatting=(patternFormatting);
                return new HSSFPatternFormatting(cfRuleRecord, workbook);
            }
            else if (Create)
            {
                patternFormatting = new PatternFormatting();
                cfRuleRecord.PatternFormatting=(patternFormatting);
                return new HSSFPatternFormatting(cfRuleRecord, workbook);
            }
            else
            {
                return null;
            }
        }

        /**
         * @return - pattern formatting object  if defined, <c>null</c> otherwise
         */
        public IPatternFormatting PatternFormatting
        {
            get { return GetPatternFormatting(false); }
        }
        /**
         * Create a new pattern formatting structure if it does not exist, 
         * otherwise just return existing object.
         * @return - pattern formatting object, never returns <c>null</c>. 
         */
        public IPatternFormatting CreatePatternFormatting()
        {
            return GetPatternFormatting(true);
        }

        private HSSFDataBarFormatting GetDataBarFormatting(bool create)
        {
            CFRule12Record cfRule12Record = GetCFRule12Record(create);
            DataBarFormatting databarFormatting = cfRule12Record.DataBarFormatting;
            if (databarFormatting != null)
            {
                return new HSSFDataBarFormatting(cfRule12Record, sheet);
            }
            else if (create)
            {
                databarFormatting = cfRule12Record.CreateDataBarFormatting();
                return new HSSFDataBarFormatting(cfRule12Record, sheet);
            }
            else
            {
                return null;
            }
        }
        /**
         * @return databar / data-bar formatting object if defined, <code>null</code> otherwise
         */
        public IDataBarFormatting DataBarFormatting
        {
            get
            {
                return GetDataBarFormatting(false);
            }
        }
        /**
         * create a new databar / data-bar formatting object if it does not exist,
         * otherwise just return the existing object.
         */
        public HSSFDataBarFormatting CreateDataBarFormatting()
        {
            return GetDataBarFormatting(true);
        }

        private HSSFIconMultiStateFormatting GetMultiStateFormatting(bool create)
        {
            CFRule12Record cfRule12Record = GetCFRule12Record(create);
            IconMultiStateFormatting iconFormatting = cfRule12Record.MultiStateFormatting;
            if (iconFormatting != null)
            {
                return new HSSFIconMultiStateFormatting(cfRule12Record, sheet);
            }
            else if (create)
            {
                iconFormatting = cfRule12Record.CreateMultiStateFormatting();
                return new HSSFIconMultiStateFormatting(cfRule12Record, sheet);
            }
            else
            {
                return null;
            }
        }

        /**
         * @return icon / multi-state formatting object if defined, <code>null</code> otherwise
         */
        public IIconMultiStateFormatting MultiStateFormatting
        {
            get { return GetMultiStateFormatting(false); }
        }

        /**
         * create a new icon / multi-state formatting object if it does not exist,
         * otherwise just return the existing object.
         */
        public HSSFIconMultiStateFormatting CreateMultiStateFormatting()
        {
            return GetMultiStateFormatting(true);
        }

        private HSSFColorScaleFormatting GetColorScaleFormatting(bool create)
        {
            CFRule12Record cfRule12Record = GetCFRule12Record(create);
            ColorGradientFormatting colorFormatting = cfRule12Record.ColorGradientFormatting;
            if (colorFormatting != null)
            {
                return new HSSFColorScaleFormatting(cfRule12Record, sheet);
            }
            else if (create)
            {
                colorFormatting = cfRule12Record.CreateColorGradientFormatting();
                return new HSSFColorScaleFormatting(cfRule12Record, sheet);
            }
            else
            {
                return null;
            }
        }
        /**
         * @return color scale / gradient formatting object if defined, <code>null</code> otherwise
         */
        public IColorScaleFormatting ColorScaleFormatting
        {
            get
            {
                return GetColorScaleFormatting(false);
            }
            
        }
        /**
         * create a new color scale / gradient formatting object if it does not exist,
         * otherwise just return the existing object.
         */
        public HSSFColorScaleFormatting CreateColorScaleFormatting()
        {
            return GetColorScaleFormatting(true);
        }
        /**
	     * @return -  the conditiontype for the cfrule
	     */
        public ConditionType ConditionType
        {
            get
            {
                return NPOI.SS.UserModel.ConditionType.ForId(cfRuleRecord.ConditionType);
            }
        }
        /**
	     * @return - the comparisionoperatation for the cfrule
	     */
        public ComparisonOperator ComparisonOperation
        {
            get
            {
                return (ComparisonOperator)cfRuleRecord.ComparisonOperation;
            }
        }

        public String Formula1
        {
            get
            {
                return ToFormulaString(cfRuleRecord.ParsedExpression1);
            }
        }

        public String Formula2
        {
            get
            {
                byte conditionType = cfRuleRecord.ConditionType;
                if (conditionType == CELL_COMPARISON)
                {
                    byte comparisonOperation = cfRuleRecord.ComparisonOperation;
                    switch (comparisonOperation)
                    {
                        case (byte)ComparisonOperator.Between:
                        case (byte)ComparisonOperator.NotBetween:
                            return ToFormulaString(cfRuleRecord.ParsedExpression2);
                    }
                }
                return null;
            }
        }

        protected internal String ToFormulaString(Ptg[] ParsedExpression)
        {
            if (ParsedExpression == null)
            {
                return null;
            }
            return HSSFConditionalFormattingRule.ToFormulaString(ParsedExpression, workbook);
        }
        protected internal static String ToFormulaString(Ptg[] parsedExpression, HSSFWorkbook workbook)
        {
            if (parsedExpression == null || parsedExpression.Length == 0)
            {
                return null;
            }
            return HSSFFormulaParser.ToFormulaString(workbook, parsedExpression);
        }
    }
}