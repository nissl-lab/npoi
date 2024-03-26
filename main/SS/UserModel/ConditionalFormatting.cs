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

    /// <summary>
    /// <para>
    /// The ConditionalFormatting class encapsulates all Settings of Conditional Formatting.
    /// </para>
    /// <para>
    /// The class can be used
    /// </para>
    /// <para>
    /// <list type="bullet">
    /// <item><description>
    /// to make a copy ConditionalFormatting Settings.
    /// </description></item>
    /// For example:
    /// <code>
    /// ConditionalFormatting cf = sheet.GetConditionalFormattingAt(index);
    /// newSheet.AddConditionalFormatting(cf);
    /// </code>
    ///  <item><description>
    ///  or to modify existing Conditional Formatting Settings (formatting regions and/or rules).
    ///  </description></item>
    ///  </list>
    /// </para>
    /// <para>
    /// Use <see cref="NPOI.HSSF.UserModel.Sheet.getSheetConditionalFormatting()" /> to Get access to an instance of this class.
    /// </para>
    /// <para>
    /// To create a new Conditional Formatting Set use the following approach:
    /// </para>
    /// <para>
    /// <code>
    /// // Define a Conditional Formatting rule, which triggers formatting
    /// // when cell's value is greater or equal than 100.0 and
    /// // applies patternFormatting defined below.
    /// ConditionalFormattingRule rule = sheet.CreateConditionalFormattingRule(
    ///     ComparisonOperator.GE,
    ///     "100.0", // 1st formula
    ///     null     // 2nd formula is not used for comparison operator GE
    /// );
    /// 
    /// // Create pattern with red background
    /// PatternFormatting patternFmt = rule.CretePatternFormatting();
    /// patternFormatting.FillBackgroundColor(IndexedColor.RED.Index);
    /// 
    /// // Define a region Containing first column
    /// Region [] regions =
    /// {
    ///     new Region(1,(short)1,-1,(short)1)
    /// };
    /// 
    /// // Apply Conditional Formatting rule defined above to the regions
    /// sheet.AddConditionalFormatting(regions, rule);
    /// </code>
    /// </para>
    /// </summary>
    /// @author Dmitriy Kumshayev
    /// @author Yegor Kozlov
    public interface IConditionalFormatting
    {

        /// <summary>
        /// </summary>
        /// <returns>array of <c>CellRangeAddress</c>s. Never <c>null</c></returns>
        CellRangeAddress[] GetFormattingRanges();
        /// <summary>
        /// Sets the cell ranges the rule conditional formatting must be applied to.
        /// </summary>
        /// <param name="ranges">non-null array of <tt>CellRangeAddress</tt>s</param>
        void SetFormattingRanges(CellRangeAddress[] ranges);
        /// <summary>
        /// Replaces an existing Conditional Formatting rule at position idx.
        /// Excel allows to create up to 3 Conditional Formatting rules.
        /// This method can be useful to modify existing  Conditional Formatting rules.
        /// </summary>
        /// <param name="idx">position of the rule. Should be between 0 and 2.</param>
        /// <param name="cfRule">- Conditional Formatting rule</param>
        void SetRule(int idx, IConditionalFormattingRule cfRule);

        /// <summary>
        /// Add a Conditional Formatting rule.
        /// Excel allows to create up to 3 Conditional Formatting rules.
        /// </summary>
        /// <param name="cfRule">- Conditional Formatting rule</param>
        void AddRule(IConditionalFormattingRule cfRule);

        /// <summary>
        /// </summary>
        /// <returns>the Conditional Formatting rule at position idx.</returns>
        IConditionalFormattingRule GetRule(int idx);

        /// <summary>
        /// </summary>
        /// <returns>number of Conditional Formatting rules.</returns>
        int NumberOfRules { get; }


    }

}