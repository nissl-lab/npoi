/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using System;
namespace NPOI.XSSF.UserModel
{

    /**
     * @author Yegor Kozlov
     */
    public class XSSFConditionalFormattingRule : IConditionalFormattingRule
    {
        private CT_CfRule _cfRule;
        private XSSFSheet _sh;

        /*package*/
        public XSSFConditionalFormattingRule(XSSFSheet sh)
        {
            _cfRule = new CT_CfRule();
            _sh = sh;
        }

        /*package*/
        internal XSSFConditionalFormattingRule(XSSFSheet sh, CT_CfRule cfRule)
        {
            _cfRule = cfRule;
            _sh = sh;
        }

        /*package*/
        internal CT_CfRule GetCTCfRule()
        {
            return _cfRule;
        }

        /*package*/
        internal CT_Dxf GetDxf(bool create)
        {
            StylesTable styles = ((XSSFWorkbook)_sh.Workbook).GetStylesSource();
            CT_Dxf dxf = null;
            if (styles.DXfsSize > 0 && _cfRule.IsSetDxfId())
            {
                int dxfId = (int)_cfRule.dxfId;
                dxf = styles.GetDxfAt(dxfId);
            }
            if (create && dxf == null)
            {
                dxf = new CT_Dxf();
                int dxfId = styles.PutDxf(dxf);
                _cfRule.dxfId = (uint)(dxfId - 1);
            }
            return dxf;
        }

        /**
         * Create a new border formatting structure if it does not exist,
         * otherwise just return existing object.
         *
         * @return - border formatting object, never returns <code>null</code>.
         */
        public IBorderFormatting CreateBorderFormatting()
        {
            CT_Dxf dxf = GetDxf(true);
            CT_Border border;
            if (!dxf.IsSetBorder())
            {
                border = dxf.AddNewBorder();
            }
            else
            {
                border = dxf.border;
            }

            return new XSSFBorderFormatting(border);
        }

        /**
         * @return - border formatting object  if defined,  <code>null</code> otherwise
         */
        public IBorderFormatting GetBorderFormatting()
        {
            CT_Dxf dxf = GetDxf(false);
            if (dxf == null || !dxf.IsSetBorder()) return null;

            return new XSSFBorderFormatting(dxf.border);
        }

        /**
         * Create a new font formatting structure if it does not exist,
         * otherwise just return existing object.
         *
         * @return - font formatting object, never returns <code>null</code>.
         */
        public IFontFormatting CreateFontFormatting()
        {
            CT_Dxf dxf = GetDxf(true);
            CT_Font font;
            if (!dxf.IsSetFont())
            {
                font = dxf.AddNewFont();
            }
            else
            {
                font = dxf.font;
            }

            return new XSSFFontFormatting(font);
        }

        /**
         * @return - font formatting object  if defined,  <code>null</code> otherwise
         */
        public IFontFormatting GetFontFormatting()
        {
            CT_Dxf dxf = GetDxf(false);
            if (dxf == null || !dxf.IsSetFont()) return null;

            return new XSSFFontFormatting(dxf.font);
        }

        /**
         * Create a new pattern formatting structure if it does not exist,
         * otherwise just return existing object.
         *
         * @return - pattern formatting object, never returns <code>null</code>.
         */
        public IPatternFormatting CreatePatternFormatting()
        {
            CT_Dxf dxf = GetDxf(true);
            CT_Fill fill;
            if (!dxf.IsSetFill())
            {
                fill = dxf.AddNewFill();
            }
            else
            {
                fill = dxf.fill;
            }

            return new XSSFPatternFormatting(fill);
        }

        /**
         * @return - pattern formatting object  if defined,  <code>null</code> otherwise
         */
        public IPatternFormatting GetPatternFormatting()
        {
            CT_Dxf dxf = GetDxf(false);
            if (dxf == null || !dxf.IsSetFill()) return null;

            return new XSSFPatternFormatting(dxf.fill);
        }

        /**
         * Type of conditional formatting rule.
         * <p>
         * MUST be either {@link ConditionalFormattingRule#CONDITION_TYPE_CELL_VALUE_IS}
         * or  {@link ConditionalFormattingRule#CONDITION_TYPE_FORMULA}
         * </p>
         *
         * @return the type of condition
         */
        public ConditionType ConditionType
        {
            get
            {
                switch (_cfRule.type)
                {
                    case ST_CfType.expression: return ConditionType.Formula;
                    case ST_CfType.cellIs: return ConditionType.CellValueIs;
                }
                return 0;
            }
        }

        /**
         * The comparison function used when the type of conditional formatting is Set to
         * {@link ConditionalFormattingRule#CONDITION_TYPE_CELL_VALUE_IS}
         * <p>
         *     MUST be a constant from {@link NPOI.ss.usermodel.ComparisonOperator}
         * </p>
         *
         * @return the conditional format operator
         */
        public ComparisonOperator ComparisonOperation
        {
            get
            {
                ST_ConditionalFormattingOperator? op = _cfRule.@operator;
                if (op == null) 
                    return ComparisonOperator.NoComparison;

                switch (op)
                {
                    case ST_ConditionalFormattingOperator.lessThan: return ComparisonOperator.LessThan;
                    case ST_ConditionalFormattingOperator.lessThanOrEqual: return ComparisonOperator.LessThanOrEqual;
                    case ST_ConditionalFormattingOperator.greaterThan: return ComparisonOperator.GreaterThan;
                    case ST_ConditionalFormattingOperator.greaterThanOrEqual: return ComparisonOperator.GreaterThanOrEqual;
                    case ST_ConditionalFormattingOperator.equal: return ComparisonOperator.Equal;
                    case ST_ConditionalFormattingOperator.notEqual: return ComparisonOperator.NotEqual;
                    case ST_ConditionalFormattingOperator.between: return ComparisonOperator.Between;
                    case ST_ConditionalFormattingOperator.notBetween: return ComparisonOperator.NotBetween;
                }
                return ComparisonOperator.NoComparison;
            }
        }

        /**
         * The formula used to Evaluate the first operand for the conditional formatting rule.
         * <p>
         * If the condition type is {@link ConditionalFormattingRule#CONDITION_TYPE_CELL_VALUE_IS},
         * this field is the first operand of the comparison.
         * If type is {@link ConditionalFormattingRule#CONDITION_TYPE_FORMULA}, this formula is used
         * to determine if the conditional formatting is applied.
         * </p>
         * <p>
         * If comparison type is {@link ConditionalFormattingRule#CONDITION_TYPE_FORMULA} the formula MUST be a Boolean function
         * </p>
         *
         * @return  the first formula
         */
        public String Formula1
        {
            get
            {
                return _cfRule.sizeOfFormulaArray() > 0 ? _cfRule.GetFormulaArray(0) : null;
            }
        }

        /**
         * The formula used to Evaluate the second operand of the comparison when
         * comparison type is  {@link ConditionalFormattingRule#CONDITION_TYPE_CELL_VALUE_IS} and operator
         * is either {@link NPOI.ss.usermodel.ComparisonOperator#BETWEEN} or {@link NPOI.ss.usermodel.ComparisonOperator#NOT_BETWEEN}
         *
         * @return  the second formula
         */
        public String Formula2
        {
            get
            {
                return _cfRule.sizeOfFormulaArray() == 2 ? _cfRule.GetFormulaArray(1) : null;
            }
        }
    }
}


