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

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS.Util;
    using NPOI.SS.UserModel;


    /// <summary>
    /// The Conditional Formatting facet of HSSFSheet
    /// @author Dmitriy Kumshayev
    /// </summary>
    public class HSSFSheetConditionalFormatting : ISheetConditionalFormatting
    {

        //private HSSFWorkbook _workbook;
        private HSSFSheet _sheet;
        private ConditionalFormattingTable _conditionalFormattingTable;

        /* package */
        //public HSSFSheetConditionalFormatting(HSSFWorkbook workbook, InternalSheet sheet)
        //{
        //    _workbook = workbook;
        //    _conditionalFormattingTable = sheet.ConditionalFormattingTable;
        //}
        public HSSFSheetConditionalFormatting(HSSFSheet sheet)
        {
            _sheet = sheet;
            _conditionalFormattingTable = sheet.Sheet.ConditionalFormattingTable;
        }
        /// <summary>
        /// A factory method allowing to Create a conditional formatting rule
        /// with a cell comparison operator
        /// TODO - formulas containing cell references are currently not Parsed properly
        /// </summary>
        /// <param name="comparisonOperation">a constant value from HSSFConditionalFormattingRule.ComparisonOperator</param>
        /// <param name="formula1">formula for the valued, Compared with the cell</param>
        /// <param name="formula2">second formula (only used with HSSFConditionalFormattingRule#COMPARISON_OPERATOR_BETWEEN 
        /// and HSSFConditionalFormattingRule#COMPARISON_OPERATOR_NOT_BETWEEN operations)</param>
        /// <returns></returns>
        public IConditionalFormattingRule CreateConditionalFormattingRule(
                ComparisonOperator comparisonOperation,
                String formula1,
                String formula2)
        {

            HSSFWorkbook wb = (HSSFWorkbook)_sheet.Workbook;
            CFRuleRecord rr = CFRuleRecord.Create(_sheet, (byte)comparisonOperation, formula1, formula2);
            return new HSSFConditionalFormattingRule(wb, rr);
        }

        /// <summary>
        /// A factory method allowing to Create a conditional formatting rule with a formula.
        /// The formatting rules are applied by Excel when the value of the formula not equal to 0.
        /// TODO - formulas containing cell references are currently not Parsed properly
        /// </summary>
        /// <param name="formula">formula for the valued, Compared with the cell</param>
        /// <returns></returns>
        public IConditionalFormattingRule CreateConditionalFormattingRule(String formula)
        {
            HSSFWorkbook wb = (HSSFWorkbook)_sheet.Workbook;
            CFRuleRecord rr = CFRuleRecord.Create(_sheet, formula);
            return new HSSFConditionalFormattingRule(wb, rr);
        }
        public IConditionalFormattingRule CreateConditionalFormattingRule(
            ComparisonOperator comparisonOperation,
            String formula1)
        {
            HSSFWorkbook wb = (HSSFWorkbook)_sheet.Workbook;
            CFRuleRecord rr = CFRuleRecord.Create(_sheet, (byte)comparisonOperation, formula1, null);
            return new HSSFConditionalFormattingRule(wb, rr);
        }
        /// <summary>
        /// Adds a copy of HSSFConditionalFormatting object to the sheet
        /// This method could be used to copy HSSFConditionalFormatting object
        /// from one sheet to another.
        /// </summary>
        /// <param name="cf">HSSFConditionalFormatting object</param>
        /// <returns>index of the new Conditional Formatting object</returns>
        /// <example>
        /// HSSFConditionalFormatting cf = sheet.GetConditionalFormattingAt(index);
        /// newSheet.AddConditionalFormatting(cf);
        /// </example>
        public int AddConditionalFormatting(IConditionalFormatting cf)
        {
            CFRecordsAggregate cfraClone = ((HSSFConditionalFormatting)cf).CFRecordsAggregate.CloneCFAggregate();

            return _conditionalFormattingTable.Add(cfraClone);
        }

        /// <summary>
        /// Allows to Add a new Conditional Formatting Set to the sheet.
        /// </summary>
        /// <param name="regions">list of rectangular regions to apply conditional formatting rules</param>
        /// <param name="cfRules">Set of up to three conditional formatting rules</param>
        /// <returns>index of the newly Created Conditional Formatting object</returns>
        public int AddConditionalFormatting(CellRangeAddress[] regions, IConditionalFormattingRule[] cfRules)
        {
            if (regions == null)
            {
                throw new ArgumentException("regions must not be null");
            }
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

            CFRuleRecord[] rules = new CFRuleRecord[cfRules.Length];
            for (int i = 0; i != cfRules.Length; i++)
            {
                rules[i] = ((HSSFConditionalFormattingRule)cfRules[i]).CfRuleRecord;
            }
            CFRecordsAggregate cfra = new CFRecordsAggregate(regions, rules);
            return _conditionalFormattingTable.Add(cfra);
        }
        public int AddConditionalFormatting(CellRangeAddress[] regions,
            HSSFConditionalFormattingRule rule1)
        {
            return AddConditionalFormatting(regions,
                    rule1 == null ? null : new HSSFConditionalFormattingRule[]
				{
					rule1
				});
        }

        /// <summary>
        /// Adds the conditional formatting.
        /// </summary>
        /// <param name="regions">The regions.</param>
        /// <param name="rule1">The rule1.</param>
        /// <returns></returns>
        public int AddConditionalFormatting(CellRangeAddress[] regions,
                IConditionalFormattingRule rule1)
        {
            return AddConditionalFormatting(regions, (HSSFConditionalFormattingRule)rule1);
        }

        /// <summary>
        /// Adds the conditional formatting.
        /// </summary>
        /// <param name="regions">The regions.</param>
        /// <param name="rule1">The rule1.</param>
        /// <param name="rule2">The rule2.</param>
        /// <returns></returns>
        public int AddConditionalFormatting(CellRangeAddress[] regions,
                IConditionalFormattingRule rule1,
                IConditionalFormattingRule rule2)
        {
            return AddConditionalFormatting(regions,
                    new HSSFConditionalFormattingRule[]
				{
						(HSSFConditionalFormattingRule)rule1, (HSSFConditionalFormattingRule)rule2
				});
        }

        /// <summary>
        /// Gets Conditional Formatting object at a particular index
        /// @param index
        /// of the Conditional Formatting object to fetch
        /// </summary>
        /// <param name="index">Conditional Formatting object</param>
        /// <returns></returns>
        public IConditionalFormatting GetConditionalFormattingAt(int index)
        {
            CFRecordsAggregate cf = _conditionalFormattingTable.Get(index);
            if (cf == null)
            {
                return null;
            }
            return new HSSFConditionalFormatting((HSSFWorkbook)_sheet.Workbook, cf);
        }

        /// <summary>
        /// the number of Conditional Formatting objects of the sheet
        /// </summary>
        /// <value>The num conditional formattings.</value>
        public int NumConditionalFormattings
        {
            get
            {
                return _conditionalFormattingTable.Count;
            }
        }

        /// <summary>
        /// Removes a Conditional Formatting object by index
        /// </summary>
        /// <param name="index">index of a Conditional Formatting object to Remove</param>
        public void RemoveConditionalFormatting(int index)
        {
            _conditionalFormattingTable.Remove(index);
        }
    }
}