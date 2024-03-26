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

    /// <summary>
    /// Represents a description of a conditional formatting rule
    /// </summary>
    /// @author Dmitriy Kumshayev
    /// @author Yegor Kozlov
    public interface IConditionalFormattingRule
    {
        /// <summary>
        /// Create a new border formatting structure if it does not exist,
        /// otherwise just return existing object.
        /// </summary>
        /// <returns>- border formatting object, never returns <c>null</c>.</returns>
        IBorderFormatting CreateBorderFormatting();

        /// <summary>
        /// </summary>
        /// <returns>- border formatting object  if defined,  <c>null</c> otherwise</returns>
        IBorderFormatting BorderFormatting { get; }

        /// <summary>
        /// Create a new font formatting structure if it does not exist,
        /// otherwise just return existing object.
        /// </summary>
        /// <returns>- font formatting object, never returns <c>null</c>.</returns>
        IFontFormatting CreateFontFormatting();

        /// <summary>
        /// </summary>
        /// <returns>- font formatting object  if defined,  <c>null</c> otherwise</returns>
        IFontFormatting FontFormatting { get; }

        /// <summary>
        /// Create a new pattern formatting structure if it does not exist,
        /// otherwise just return existing object.
        /// </summary>
        /// <returns>- pattern formatting object, never returns <c>null</c>.</returns>
        IPatternFormatting CreatePatternFormatting();

        /// <summary>
        /// </summary>
        /// <returns>- pattern formatting object  if defined,  <c>null</c> otherwise</returns>
        IPatternFormatting PatternFormatting { get; }


        /// <summary>
        /// </summary>
        /// <returns>- databar / data-bar formatting object if defined, <c>null</c> otherwise</returns>
        IDataBarFormatting DataBarFormatting { get; }

        /// <summary>
        /// </summary>
        /// <returns>- icon / multi-state formatting object if defined, <c>null</c> otherwise</returns>
        IIconMultiStateFormatting MultiStateFormatting { get; }

        /// <summary>
        /// </summary>
        /// <returns>color scale / color grate formatting object if defined, <c>null</c> otherwise</returns>
        IColorScaleFormatting ColorScaleFormatting { get; }

        /// <summary>
        /// <para>
        /// Type of conditional formatting rule.
        /// </para>
        /// <para>
        /// MUST be either <see cref="CONDITION_TYPE_CELL_VALUE_IS"/> or  <see cref="CONDITION_TYPE_FORMULA"/>
        /// </para>
        /// </summary>
        /// <returns>the type of condition</returns>
        ConditionType ConditionType { get; }

        /// <summary>
        /// <para>
        /// The comparison function used when the type of conditional formatting is Set to
        /// <see cref="CONDITION_TYPE_CELL_VALUE_IS"/>
        /// </para>
        /// <para>
        ///     MUST be a constant from <see cref="ComparisonOperator"/>
        /// </para>
        /// </summary>
        /// <returns>the conditional format operator</returns>
        ComparisonOperator ComparisonOperation { get; }

        /// <summary>
        /// <para>
        /// The formula used to Evaluate the first operand for the conditional formatting rule.
        /// </para>
        /// <para>
        /// If the condition type is <see cref="CONDITION_TYPE_CELL_VALUE_IS"/>,
        /// this field is the first operand of the comparison.
        /// If type is <see cref="CONDITION_TYPE_FORMULA"/>, this formula is used
        /// to determine if the conditional formatting is applied.
        /// </para>
        /// <para>
        /// 
        /// </para>
        /// <para>
        /// If comparison type is <see cref="CONDITION_TYPE_FORMULA"/> the formula MUST be a Boolean function
        /// </para>
        /// </summary>
        /// <returns>the first formula</returns>
        String Formula1 { get; }

        /// <summary>
        /// The formula used to Evaluate the second operand of the comparison when
        /// comparison type is  <see cref="CONDITION_TYPE_CELL_VALUE_IS"/> and operator
        /// is either <see cref="ComparisonOperator.BETWEEN" /> or <see cref="ComparisonOperator.NOT_BETWEEN" />
        /// </summary>
        /// <returns>the second formula</returns>
        String Formula2 { get; }

        /// <summary>
        /// XSSF rules store textual condition values as an attribute and also as a formula that needs shifting.  Using the attribute is simpler/faster.
        /// HSSF rules don't have this and return null.  We can fall back on the formula for those (AFAIK).
        /// @return condition text if it exists, or null
        /// </summary>
        String Text { get; }

        /// <summary>
        /// The priority of the rule, if defined, otherwise 0.
        /// For XSSF, this should always be set.For HSSF, only newer style rules
        /// have this set, older ones will return 0.
        /// </summary>
        /// <returns></returns>
        int Priority { get; }

        bool StopIfTrue { get; }

        ExcelNumberFormat NumberFormat { get; }

        ConditionFilterType? ConditionFilterType { get; }

        IConditionFilterData FilterConfiguration { get; }
    }

}