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

namespace NPOI.SS.UserModel
{
    using System;

    using NPOI.SS.Util;

    /**
     * The 'Conditional Formatting' facet of <c>Sheet</c>
     *
     * @author Dmitriy Kumshayev
     * @author Yegor Kozlov
     * @since 3.8
     */
    public interface ISheetConditionalFormatting
    {
        /// <summary>
        /// Add a new Conditional Formatting to the sheet.
        /// </summary>
        /// <param name="regions">list of rectangular regions to apply conditional formatting rules</param>
        /// <param name="rule">the rule to apply</param>
        /// <returns>index of the newly Created Conditional Formatting object</returns>
        int AddConditionalFormatting(CellRangeAddress[] regions,
                                     IConditionalFormattingRule rule);

        /// <summary>
        /// Add a new Conditional Formatting consisting of two rules.
        /// </summary>
        /// <param name="regions"> list of rectangular regions to apply conditional formatting rules</param>
        /// <param name="rule1">the first rule</param>
        /// <param name="rule2">the second rule</param>
        /// <returns>index of the newly Created Conditional Formatting object</returns>
        int AddConditionalFormatting(CellRangeAddress[] regions,
                                     IConditionalFormattingRule rule1,
                                     IConditionalFormattingRule rule2);

        /// <summary>
        /// Add a new Conditional Formatting Set to the sheet.
        /// </summary>
        /// <param name="regions">list of rectangular regions to apply conditional formatting rules</param>
        /// <param name="cfRules">Set of up to three conditional formatting rules</param>
        /// <returns>index of the newly Created Conditional Formatting object</returns>
        int AddConditionalFormatting(CellRangeAddress[] regions, IConditionalFormattingRule[] cfRules);

        /// <summary>
        /// Adds a copy of a ConditionalFormatting object to the sheet
        /// </summary>
        /// <param name="cf">the Conditional Formatting to clone</param>
        /// <returns>index of the new Conditional Formatting object</returns>
        /// <remarks>
        /// This method could be used to copy ConditionalFormatting object
        /// from one sheet to another. For example:
        /// ConditionalFormatting cf = sheet.GetConditionalFormattingAt(index);
        /// newSheet.AddConditionalFormatting(cf);
        /// </remarks>
        int AddConditionalFormatting(IConditionalFormatting cf);

        /**
         * A factory method allowing to create a conditional formatting rule
         * with a cell comparison operator
         * <p>
         * The Created conditional formatting rule Compares a cell value
         * to a formula calculated result, using the specified operator.
         * The type  of the Created condition is {@link ConditionalFormattingRule#CONDITION_TYPE_CELL_VALUE_IS}
         * </p>
         *
         * @param comparisonOperation - MUST be a constant value from
         *		 <c>{@link ComparisonOperator}</c>: <p>
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
         * {@link ComparisonOperator#BETWEEN}) and {@link ComparisonOperator#NOT_BETWEEN} operations)
         */
        IConditionalFormattingRule CreateConditionalFormattingRule(
                ComparisonOperator comparisonOperation,
                String formula1,
                String formula2);


        /// <summary>
        /// Create a conditional formatting rule that Compares a cell value to a formula calculated result, using an operator 
        /// </summary>
        /// <param name="comparisonOperation">MUST be a constant value from ComparisonOperator except  BETWEEN and NOT_BETWEEN</param>
        /// <param name="formula">the formula to determine if the conditional formatting is applied</param>
        /// <returns>a conditional formatting rule</returns>
        IConditionalFormattingRule CreateConditionalFormattingRule(
                ComparisonOperator comparisonOperation,
                String formula);

        /// <summary>
        /// Create a conditional formatting rule based on a Boolean formula.
        /// When the formula result is true, the cell is highlighted.
        /// </summary>
        /// <param name="formula">the formula to Evaluate. MUST be a Boolean function.</param>
        /// <returns>conditional formatting rule</returns>
        IConditionalFormattingRule CreateConditionalFormattingRule(String formula);

        /// <summary>
        /// Gets Conditional Formatting object at a particular index
        /// </summary>
        /// <param name="index">0-based index of the Conditional Formatting object to fetch</param>
        /// <returns>Conditional Formatting object or null if not found</returns>
        /// <exception cref="ArgumentException">throws ArgumentException if the index is  outside of the allowable range (0 ... numberOfFormats-1)</exception>
        IConditionalFormatting GetConditionalFormattingAt(int index);

        /// <summary>
        /// get the number of conditional formats in this sheet
        /// </summary>
        int NumConditionalFormattings { get; }

        /// <summary>
        /// Removes a Conditional Formatting object by index
        /// </summary>
        /// <param name="index">0-based index of the Conditional Formatting object to remove</param>
        /// <exception cref="ArgumentException">throws ArgumentException if the index is  outside of the allowable range (0 ... numberOfFormats-1)</exception>
        void RemoveConditionalFormatting(int index);
    }

}