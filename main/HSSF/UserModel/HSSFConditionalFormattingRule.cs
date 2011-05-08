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
    using System.Text;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record.CF;
    using NPOI.SS.Formula;

    /**
     * 
     * High level representation of Conditional Formatting Rule.
     * It allows to specify formula based conditions for the Conditional Formatting
     * and the formatting Settings such as font, border and pattern.
     * 
     * @author Dmitriy Kumshayev
     */

    public class HSSFConditionalFormattingRule
    {
        private static byte CELL_COMPARISON = CFRuleRecord.CONDITION_TYPE_CELL_VALUE_IS;

        private CFRuleRecord cfRuleRecord;
        private HSSFWorkbook workbook;

        public HSSFConditionalFormattingRule(HSSFWorkbook pWorkbook, CFRuleRecord pRuleRecord)
        {
            workbook = pWorkbook;
            cfRuleRecord = pRuleRecord;
        }

        public CFRuleRecord CfRuleRecord
        {
            get { return cfRuleRecord; }
        }

        private HSSFFontFormatting GetFontFormatting(bool Create)
        {
            FontFormatting fontFormatting = cfRuleRecord.FontFormatting;
            if (fontFormatting != null)
            {
                cfRuleRecord.FontFormatting=(fontFormatting);
                return new HSSFFontFormatting(cfRuleRecord);
            }
            else if (Create)
            {
                fontFormatting = new FontFormatting();
                cfRuleRecord.FontFormatting=(fontFormatting);
                return new HSSFFontFormatting(cfRuleRecord);
            }
            else
            {
                return null;
            }
        }

        /**
         * @return - font formatting object  if defined,  <c>null</c> otherwise
         */
        public HSSFFontFormatting GetFontFormatting()
        {
            return GetFontFormatting(false);
        }
        /**
         * Create a new font formatting structure if it does not exist, 
         * otherwise just return existing object.
         * @return - font formatting object, never returns <c>null</c>. 
         */
        public HSSFFontFormatting CreateFontFormatting()
        {
            return GetFontFormatting(true);
        }

        private HSSFBorderFormatting GetBorderFormatting(bool Create)
        {
            BorderFormatting borderFormatting = cfRuleRecord.BorderFormatting;
            if (borderFormatting != null)
            {
                cfRuleRecord.BorderFormatting=(borderFormatting);
                return new HSSFBorderFormatting(cfRuleRecord);
            }
            else if (Create)
            {
                borderFormatting = new BorderFormatting();
                cfRuleRecord.BorderFormatting=(borderFormatting);
                return new HSSFBorderFormatting(cfRuleRecord);
            }
            else
            {
                return null;
            }
        }
        /**
         * @return - border formatting object  if defined,  <c>null</c> otherwise
         */
        public HSSFBorderFormatting GetBorderFormatting()
        {
            return GetBorderFormatting(false);
        }
        /**
         * Create a new border formatting structure if it does not exist, 
         * otherwise just return existing object.
         * @return - border formatting object, never returns <c>null</c>. 
         */
        public HSSFBorderFormatting CreateBorderFormatting()
        {
            return GetBorderFormatting(true);
        }

        private HSSFPatternFormatting GetPatternFormatting(bool Create)
        {
            PatternFormatting patternFormatting = cfRuleRecord.PatternFormatting;
            if (patternFormatting != null)
            {
                cfRuleRecord.PatternFormatting=(patternFormatting);
                return new HSSFPatternFormatting(cfRuleRecord);
            }
            else if (Create)
            {
                patternFormatting = new PatternFormatting();
                cfRuleRecord.PatternFormatting=(patternFormatting);
                return new HSSFPatternFormatting(cfRuleRecord);
            }
            else
            {
                return null;
            }
        }

        /**
         * @return - pattern formatting object  if defined, <c>null</c> otherwise
         */
        public HSSFPatternFormatting GetPatternFormatting()
        {
            return GetPatternFormatting(false);
        }
        /**
         * Create a new pattern formatting structure if it does not exist, 
         * otherwise just return existing object.
         * @return - pattern formatting object, never returns <c>null</c>. 
         */
        public HSSFPatternFormatting CreatePatternFormatting()
        {
            return GetPatternFormatting(true);
        }

        public String Formula1
        {
            get
            {
                return ToFormulaString(cfRuleRecord.GetParsedExpression1());
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
                        case (byte)ComparisonOperator.BETWEEN:
                        case (byte)ComparisonOperator.NOT_BETWEEN:
                            return ToFormulaString(cfRuleRecord.ParsedExpression2);
                    }
                }
                return null;
            }
        }

        private String ToFormulaString(Ptg[] ParsedExpression)
        {
            if (ParsedExpression == null)
            {
                return null;
            }
            return HSSFFormulaParser.ToFormulaString(workbook, ParsedExpression);
        }
    }
}