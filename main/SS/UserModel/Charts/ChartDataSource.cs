/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
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

namespace NPOI.SS.UserModel.Charts
{
    public interface IChartDataSource<T>
    {
        /// <summary>
        /// Return number of points contained by data source.
        /// </summary>
        /// <returns>number of points contained by data source</returns>
        int PointCount { get; }

        /// <summary>
        /// Returns point value at specified index.
        /// </summary>
        /// <param name="index">index to value from</param>
        /// <returns>point value at specified index.</returns>
        /// <exception cref="{@code">IndexOutOfBoundsException} if index
        /// parameter not in range {@code 0 &lt;= index &lt;= pointCount}
        /// </exception>
        T GetPointAt(int index);

        /// <summary>
        /// Returns <c>true</c> if charts data source is valid cell range.
        /// </summary>
        /// <returns><c>true</c> if charts data source is valid cell range</returns>
        bool IsReference { get; }

        /// <summary>
        /// Returns <c>true</c> if data source points should be treated as numbers.
        /// </summary>
        /// <returns><c>true</c> if data source points should be treated as numbers</returns>
        bool IsNumeric { get; }

        /// <summary>
        /// Returns formula representation of the data source. It is only applicable
        /// for data source that is valid cell range.
        /// </summary>
        /// <returns>formula representation of the data source</returns>
        /// <exception cref="{@code">UnsupportedOperationException} if the data source is not a
        /// reference.
        /// </exception>
        string FormulaString { get; }
    }
}