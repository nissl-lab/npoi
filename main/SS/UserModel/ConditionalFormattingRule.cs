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
    public interface IConditionalFormattingRule : IDifferentialStyleProvider
    {
        /// <summary>
        /// Create a new border formatting structure if it does not exist,
        /// otherwise just return existing object.
        /// </summary>
        /// <return>- border formatting object, never returns <c>null</c>.</return>
        IBorderFormatting CreateBorderFormatting();

        /// <summary>
        /// border formatting object  if defined,  <c>null</c> otherwise
        /// </summary>
        IBorderFormatting BorderFormatting { get; }

        /// <summary>
        /// Create a new font formatting structure if it does not exist,
        /// otherwise just return existing object.
        /// </summary>
        /// <return>- font formatting object, never returns <c>null</c>.</return>
        IFontFormatting CreateFontFormatting();

        /// <summary>
        /// font formatting object  if defined,  <c>null</c> otherwise
        /// </summary>
        IFontFormatting FontFormatting { get; }

        /// <summary>
        /// Create a new pattern formatting structure if it does not exist,
        /// otherwise just return existing object.
        /// </summary>
        /// <return>- pattern formatting object, never returns <c>null</c>.</return>
        IPatternFormatting CreatePatternFormatting();

        /// <summary>
        /// pattern formatting object  if defined,  <c>null</c> otherwise
        /// </summary>
        IPatternFormatting PatternFormatting { get; }


        /// <summary>
        /// databar / data-bar formatting object if defined, <c>null</c> otherwise
        /// </summary>
        IDataBarFormatting DataBarFormatting { get; }

        /// <summary>
        /// icon / multi-state formatting object if defined, <c>null</c> otherwise
        /// </summary>
        IIconMultiStateFormatting MultiStateFormatting { get; }

        /// <summary>
        /// color scale / color grate formatting object if defined, <c>null</c> otherwise
        /// </summary>
        IColorScaleFormatting ColorScaleFormatting { get; }

        /// <summary>
        /// Type of conditional formatting rule.
        /// </summary>
        ConditionType ConditionType { get; }

        /// <summary>
        /// <para>
        /// The comparison function used when the type of conditional formatting is Set to
        /// <see cref="ConditionType.CELL_VALUE_IS" />
        /// </para>
        /// <para>
        /// MUST be a constant from <see cref="ComparisonOperator"/>
        /// </para>
        /// </summary>
        /// <return>the conditional format operator</return>
        ComparisonOperator ComparisonOperation { get; }

        /// <summary>
        /// <para>
        /// The formula used to evaluate the first operand for the conditional formatting rule.
        /// </para>
        /// <para>
        /// If the condition type is <see cref="ConditionType.CELL_VALUE_IS" />,
        /// this field is the first operand of the comparison.
        /// If type is <see cref="ConditionType.FORMULA" />, this formula is used
        /// to determine if the conditional formatting is applied.
        /// </para>
        /// <para>
        /// If comparison type is <see cref="ConditionType.FORMULA" /> the formula MUST be a Boolean function
        /// </para>
        /// </summary>
        /// <return>the first formula</return>
        String Formula1 { get; }

        /// <summary>
        /// The formula used to evaluate the second operand of the comparison when
        /// comparison type is  <see cref="ConditionType.CELL_VALUE_IS" /> and operator
        /// is either <see cref="ComparisonOperator.BETWEEN" /> or <see cref="ComparisonOperator.NOT_BETWEEN" />
        /// </summary>
        /// <return>the second formula</return>
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
        int Priority {get; }

        /// <summary>
        /// <para>
        /// Always true for HSSF rules, optional flag for XSSF rules.
        /// See Excel help for more.
        /// </para>
        /// <para>
        /// true if conditional formatting rule processing stops when this one is true, false if not
        /// </para>
        /// </summary>
        /// <see cref="https://support.office.com/en-us/article/Manage-conditional-formatting-rule-precedence-063cde21-516e-45ca-83f5-8e8126076249" /> Microsoft Excel help
        bool StopIfTrue { get; }
        
        /// <summary>
        /// number format defined for this rule, or null if the cell default should be used
        /// </summary>
        ExcelNumberFormat NumberFormat { get; }

        /// <summary>
        /// This is null if
        /// <p/>
        /// <code><see cref="getConditionType()" /> != <see cref="ConditionType.FILTER" /></code>
        /// <p/>
        /// This is always <see cref="ConditionFilterType.FILTER" /> for HSSF rules of type <see cref="ConditionType.FILTER" />.
        /// <p/>
        /// For XSSF filter rules, this will indicate the specific type of filter.
        /// </summary>
        /// <return>filter type for filter rules, or null if not a filter rule.</return>
        ConditionFilterType? ConditionFilterType { get; }

        /// <summary>
        /// This is null if
        /// <p/>
        /// <code><see cref="getConditionFilterType()" /> == null</code>
        /// <p/>
        /// This means it is always null for HSSF, which does not define the extended condition types.
        /// <p/>
        /// This object contains the additional configuration information for XSSF filter conditions.
        /// </summary>
        /// <return>the Filter Configuration Data, or null if there isn't any</return>
        IConditionFilterData FilterConfiguration { get; }
    }

}