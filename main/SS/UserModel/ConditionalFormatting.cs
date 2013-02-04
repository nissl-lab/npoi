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
    using NPOI.SS.Util;

    /**
     * The ConditionalFormatting class encapsulates all Settings of Conditional Formatting.
     *
     * The class can be used
     *
     * <UL>
     * <LI>
     * to make a copy ConditionalFormatting Settings.
     * </LI>
     *
     *
     * For example:
     * <PRE>
     * ConditionalFormatting cf = sheet.GetConditionalFormattingAt(index);
     * newSheet.AddConditionalFormatting(cf);
     * </PRE>
     *
     *  <LI>
     *  or to modify existing Conditional Formatting Settings (formatting regions and/or rules).
     *  </LI>
     *  </UL>
     *
     * Use {@link NPOI.HSSF.UserModel.Sheet#getSheetConditionalFormatting()} to Get access to an instance of this class.
     * 
     * To create a new Conditional Formatting Set use the following approach:
     *
     * <PRE>
     *
     * // Define a Conditional Formatting rule, which triggers formatting
     * // when cell's value is greater or equal than 100.0 and
     * // applies patternFormatting defined below.
     * ConditionalFormattingRule rule = sheet.CreateConditionalFormattingRule(
     *     ComparisonOperator.GE,
     *     "100.0", // 1st formula
     *     null     // 2nd formula is not used for comparison operator GE
     * );
     *
     * // Create pattern with red background
     * PatternFormatting patternFmt = rule.CretePatternFormatting();
     * patternFormatting.FillBackgroundColor(IndexedColor.RED.Index);
     *
     * // Define a region Containing first column
     * Region [] regions =
     * {
     *     new Region(1,(short)1,-1,(short)1)
     * };
     *
     * // Apply Conditional Formatting rule defined above to the regions
     * sheet.AddConditionalFormatting(regions, rule);
     * </PRE>
     *
     * @author Dmitriy Kumshayev
     * @author Yegor Kozlov
     */
    public interface IConditionalFormatting
    {

        /**
         * @return array of <c>CellRangeAddress</c>s. Never <code>null</code>
         */
        CellRangeAddress[] GetFormattingRanges();

        /**
         * Replaces an existing Conditional Formatting rule at position idx.
         * Excel allows to create up to 3 Conditional Formatting rules.
         * This method can be useful to modify existing  Conditional Formatting rules.
         *
         * @param idx position of the rule. Should be between 0 and 2.
         * @param cfRule - Conditional Formatting rule
         */
        void SetRule(int idx, IConditionalFormattingRule cfRule);

        /**
         * Add a Conditional Formatting rule.
         * Excel allows to create up to 3 Conditional Formatting rules.
         *
         * @param cfRule - Conditional Formatting rule
         */
        void AddRule(IConditionalFormattingRule cfRule);

        /**
         * @return the Conditional Formatting rule at position idx.
         */
        IConditionalFormattingRule GetRule(int idx);

        /**
         * @return number of Conditional Formatting rules.
         */
        int NumberOfRules { get; }


    }

}