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

using System;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Record.CF;
using System.Collections.Generic;
using NPOI.SS;
namespace NPOI.XSSF.UserModel
{





    /**
     * @author Yegor Kozlov
     */
    public class XSSFSheetConditionalFormatting : ISheetConditionalFormatting
    {
        private XSSFSheet _sheet;

        /* namespace */
        internal XSSFSheetConditionalFormatting(XSSFSheet sheet)
        {
            _sheet = sheet;
        }

        /**
         * A factory method allowing to create a conditional formatting rule
         * with a cell comparison operator<p/>
         * TODO - formulas Containing cell references are currently not Parsed properly
         *
         * @param comparisonOperation - a constant value from
         *		 <tt>{@link NPOI.hssf.record.CFRuleRecord.ComparisonOperator}</tt>: <p>
         * <ul>
         *		 <li>BETWEEN</li>
         *		 <li>NOT_BETWEEN</li>
         *		 <li>EQUAL</li>
         *		 <li>NOT_EQUAL</li>
         *		 <li>GT</li>
         *		 <li>LT</li>
         *		 <li>GE</li>
         *		 <li>LE</li>
         * </ul>
         * </p>
         * @param formula1 - formula for the valued, Compared with the cell
         * @param formula2 - second formula (only used with
         * {@link NPOI.ss.usermodel.ComparisonOperator#BETWEEN}) and
         * {@link NPOI.ss.usermodel.ComparisonOperator#NOT_BETWEEN} operations)
         */
        public IConditionalFormattingRule CreateConditionalFormattingRule(
                ComparisonOperator comparisonOperation,
                String formula1,
                String formula2)
        {

            XSSFConditionalFormattingRule rule = new XSSFConditionalFormattingRule(_sheet);
            CT_CfRule cfRule = rule.GetCTCfRule();
            cfRule.AddFormula(formula1);
            if (formula2 != null) cfRule.AddFormula(formula2);
            cfRule.type = (ST_CfType.cellIs);
            ST_ConditionalFormattingOperator operator1;
            switch (comparisonOperation)
            {
                case ComparisonOperator.Between: 
                    operator1 = ST_ConditionalFormattingOperator.between; break;
                case ComparisonOperator.NotBetween: 
                    operator1 = ST_ConditionalFormattingOperator.notBetween; break;
                case ComparisonOperator.LessThan: 
                    operator1 = ST_ConditionalFormattingOperator.lessThan; break;
                case ComparisonOperator.LessThanOrEqual: 
                    operator1 = ST_ConditionalFormattingOperator.lessThanOrEqual; break;
                case ComparisonOperator.GreaterThan: 
                    operator1 = ST_ConditionalFormattingOperator.greaterThan; break;
                case ComparisonOperator.GreaterThanOrEqual: 
                    operator1 = ST_ConditionalFormattingOperator.greaterThanOrEqual; break;
                case ComparisonOperator.Equal: 
                    operator1 = ST_ConditionalFormattingOperator.equal; break;
                case ComparisonOperator.NotEqual: 
                    operator1 = ST_ConditionalFormattingOperator.notEqual; break;
                default: 
                    throw new ArgumentException("Unknown comparison operator: " + comparisonOperation);
            }
            cfRule.@operator = (operator1);

            return rule;
        }

        public IConditionalFormattingRule CreateConditionalFormattingRule(
                ComparisonOperator comparisonOperation,
                String formula)
        {

            return CreateConditionalFormattingRule(comparisonOperation, formula, null);
        }

        /**
         * A factory method allowing to create a conditional formatting rule with a formula.<br>
         *
         * @param formula - formula for the valued, Compared with the cell
         */
        public IConditionalFormattingRule CreateConditionalFormattingRule(String formula)
        {
            XSSFConditionalFormattingRule rule = new XSSFConditionalFormattingRule(_sheet);
            CT_CfRule cfRule = rule.GetCTCfRule();
            cfRule.AddFormula(formula);
            cfRule.type = ST_CfType.expression;
            return rule;
        }

        public int AddConditionalFormatting(CellRangeAddress[] regions, IConditionalFormattingRule[] cfRules)
        {
            if (regions == null)
            {
                throw new ArgumentException("regions must not be null");
            }
            foreach (CellRangeAddress range in regions) range.Validate(SpreadsheetVersion.EXCEL2007);

            if (cfRules == null)
            {
                throw new ArgumentException("cfRules must not be null");
            }
            if (cfRules.Length == 0)
            {
                throw new ArgumentException("cfRules must not be empty");
            }
            if (cfRules.Length > 3)
            {
                throw new ArgumentException("Number of rules must not exceed 3");
            }

            CellRangeAddress[] mergeCellRanges = CellRangeUtil.MergeCellRanges(regions);
            CT_ConditionalFormatting cf = _sheet.GetCTWorksheet().AddNewConditionalFormatting();
            string refs = string.Empty;
            foreach (CellRangeAddress a in mergeCellRanges)
            {
                if (refs.Length == 0)
                    refs = a.FormatAsString();
                else
                    refs += " " +a.FormatAsString() ;
            }
            cf.sqref = refs;

            int priority = 1;
            foreach (CT_ConditionalFormatting c in _sheet.GetCTWorksheet().conditionalFormatting)
            {
                priority += c.sizeOfCfRuleArray();
            }

            foreach (IConditionalFormattingRule rule in cfRules)
            {
                XSSFConditionalFormattingRule xRule = (XSSFConditionalFormattingRule)rule;
                xRule.GetCTCfRule().priority = (priority++);
                cf.AddNewCfRule().Set(xRule.GetCTCfRule());
            }
            return _sheet.GetCTWorksheet().SizeOfConditionalFormattingArray() - 1;
        }

        public int AddConditionalFormatting(CellRangeAddress[] regions,
                IConditionalFormattingRule rule1)
        {
            return AddConditionalFormatting(regions,
                    rule1 == null ? null : new XSSFConditionalFormattingRule[] {
                (XSSFConditionalFormattingRule)rule1
        });
        }

        public int AddConditionalFormatting(CellRangeAddress[] regions,
                IConditionalFormattingRule rule1, IConditionalFormattingRule rule2)
        {
            return AddConditionalFormatting(regions,
                    rule1 == null ? null : new XSSFConditionalFormattingRule[] {
                    (XSSFConditionalFormattingRule)rule1,
                    (XSSFConditionalFormattingRule)rule2
            });
        }

        /**
         * Adds a copy of HSSFConditionalFormatting object to the sheet
         * <p>This method could be used to copy HSSFConditionalFormatting object
         * from one sheet to another. For example:
         * <pre>
         * HSSFConditionalFormatting cf = sheet.GetConditionalFormattingAt(index);
         * newSheet.AddConditionalFormatting(cf);
         * </pre>
         *
         * @param cf HSSFConditionalFormatting object
         * @return index of the new Conditional Formatting object
         */
        public int AddConditionalFormatting(IConditionalFormatting cf)
        {
            XSSFConditionalFormatting xcf = (XSSFConditionalFormatting)cf;
            CT_Worksheet sh = _sheet.GetCTWorksheet();
            sh.AddNewConditionalFormatting().Set(xcf.GetCTConditionalFormatting());//this is already copied in Set -> .Copy()); ommitted
            return sh.SizeOfConditionalFormattingArray() - 1;
        }

        /**
        * Gets Conditional Formatting object at a particular index
        *
        * @param index
        *			of the Conditional Formatting object to fetch
        * @return Conditional Formatting object
        */
        public IConditionalFormatting GetConditionalFormattingAt(int index)
        {
            CheckIndex(index);
            CT_ConditionalFormatting cf = _sheet.GetCTWorksheet().GetConditionalFormattingArray(index);
            return new XSSFConditionalFormatting(_sheet, cf);
        }

        /**
        * @return number of Conditional Formatting objects of the sheet
        */
        public int NumConditionalFormattings
        {
            get
            {
                return _sheet.GetCTWorksheet().SizeOfConditionalFormattingArray();
            }
        }

        /**
        * Removes a Conditional Formatting object by index
        * @param index of a Conditional Formatting object to remove
        */
        public void RemoveConditionalFormatting(int index)
        {
            CheckIndex(index);
            _sheet.GetCTWorksheet().conditionalFormatting.RemoveAt(index);
        }

        private void CheckIndex(int index)
        {
            int cnt = NumConditionalFormattings;
            if (index < 0 || index >= cnt)
            {
                throw new ArgumentException("Specified CF index " + index
                        + " is outside the allowable range (0.." + (cnt - 1) + ")");
            }
        }

    }

}
