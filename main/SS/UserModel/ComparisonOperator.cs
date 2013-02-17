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
    /**
     * The conditional format operators used for "Highlight Cells That Contain..." rules.
     * <p>
     * For example, "highlight cells that begin with "M2" and contain "Mountain Gear".
     * </p>
     *
     * @author Dmitriy Kumshayev
     * @author Yegor Kozlov
     */
    public enum ComparisonOperator:byte
    {
        NoComparison = 0,

        /**
         * 'Between' operator
         */
        Between = 1,

        /**
         * 'Not between' operator
         */
        NotBetween = 2,

        /**
         *  'Equal to' operator
         */
        Equal = 3,

        /**
         * 'Not equal to' operator
         */
        NotEqual = 4,

        /**
         * 'Greater than' operator
         */
        GreaterThan = 5,

        /**
         * 'Less than' operator
         */
        LessThan = 6,

        /**
         * 'Greater than or equal to' operator
         */
        GreaterThanOrEqual = 7,

        /**
         * 'Less than or equal to' operator
         */
        LessThanOrEqual = 8,
    }

}