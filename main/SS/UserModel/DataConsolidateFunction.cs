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

    /**
     * Enum mapping the values of STDataConsolidateFunction
     */
    public class DataConsolidateFunction
    {
        public static DataConsolidateFunction AVERAGE = new DataConsolidateFunction(1, "Average"),
        COUNT = new DataConsolidateFunction(2, "Count"),
        COUNT_NUMS = new DataConsolidateFunction(3, "Count"),
        MAX = new DataConsolidateFunction(4, "Max"),
        MIN = new DataConsolidateFunction(5, "Min"),
        PRODUCT = new DataConsolidateFunction(6, "Product"),
        STD_DEV = new DataConsolidateFunction(7, "StdDev"),
        STD_DEVP = new DataConsolidateFunction(8, "StdDevp"),
        SUM = new DataConsolidateFunction(9, "Sum"),
        VAR = new DataConsolidateFunction(10, "Var"),
        VARP = new DataConsolidateFunction(11, "Varp");

        private int value;
        private String name;

        public DataConsolidateFunction(int value, String name)
        {
            this.value = value;
            this.name = name;
        }

        public String Name
        {
            get
            {
                return this.name;
            }
        }

        public int Value
        {
            get
            {
                return this.value;
            }
        }
    }
}