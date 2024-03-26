/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"), you may not use this file except in compliance with
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
    /// <summary>
    /// <para>
    /// The conditional format operators used for "Highlight Cells That Contain..." rules.
    /// </para>
    /// <para>
    /// For example, "highlight cells that begin with "M2" and contain "Mountain Gear".
    /// </para>
    /// </summary>
    /// @author Dmitriy Kumshayev
    /// @author Yegor Kozlov
    public enum ComparisonOperator : byte
    {
        NoComparison = 0,

        /// <summary>
        /// 'Between' operator
        /// </summary>
        Between = 1,

        /// <summary>
        /// 'Not between' operator
        /// </summary>
        NotBetween = 2,

        /// <summary>
        ///  'Equal to' operator
        /// </summary>
        Equal = 3,

        /// <summary>
        /// 'Not equal to' operator
        /// </summary>
        NotEqual = 4,

        /// <summary>
        /// 'Greater than' operator
        /// </summary>
        GreaterThan = 5,

        /// <summary>
        /// 'Less than' operator
        /// </summary>
        LessThan = 6,

        /// <summary>
        /// 'Greater than or equal to' operator
        /// </summary>
        GreaterThanOrEqual = 7,

        /// <summary>
        /// 'Less than or equal to' operator
        /// </summary>
        LessThanOrEqual = 8,
    }
}