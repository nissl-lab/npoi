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

using System.Collections.Generic;

namespace NPOI.SS.UserModel
{
    using System;


    /**
     * The Threshold / CFVO / Conditional Formatting Value Object.
     * <p>This defines how to calculate the ranges for a conditional
     *  formatting rule, eg which values Get a Green Traffic Light
     *  icon and which Yellow or Red.</p>
     */
    public interface IConditionalFormattingThreshold {



        /**
         * Get the Range Type used
         */
        RangeType RangeType { get; set; }

        /**
         * Changes the Range Type used
         * 
         * <p>If you change the range type, you need to
         *  ensure that the Formula and Value parameters
         *  are compatible with it before saving</p>
         */


        /**
         * Formula to use to calculate the threshold,
         *  or <code>null</code> if no formula 
         */
        String Formula { get; set; }

        /**
         * Sets the formula used to calculate the threshold,
         *  or unsets it if <code>null</code> is given.
         */

        /**
         * Gets the value used for the threshold, or 
         *  <code>null</code> if there isn't one.
         */
        double? Value { get; set; }

        /**
         * Sets the value used for the threshold. 
         * <p>If the type is {@link RangeType#PERCENT} or 
         *  {@link RangeType#PERCENTILE} it must be between 0 and 100.
         * <p>If the type is {@link RangeType#MIN} or {@link RangeType#MAX}
         *  or {@link RangeType#FORMULA} it shouldn't be Set.
         * <p>Use <code>null</code> to unset
         */
    }
    public class RangeType
    {
        /** Number / Parameter */
        public static RangeType NUMBER = new RangeType(1, "num");
        /** The minimum value from the range */
        public static RangeType MIN = new RangeType(2, "min");
        /** The maximum value from the range */
        public static RangeType MAX = new RangeType(3, "max");
        /** Percent of the way from the mi to the max value in the range */
        public static RangeType PERCENT = new RangeType(4, "percent");
        /** The minimum value of the cell that is in X percentile of the range */
        public static RangeType PERCENTILE = new RangeType(5, "percentile");
        public static RangeType UNALLOCATED = new RangeType(6, null);
        /** Formula result */
        public static RangeType FORMULA = new RangeType(7, "formula");

        public static RangeType AUTOMIN = new RangeType(8, "autoMin");

        public static RangeType AUTOMAX = new RangeType(9, "autoMax");

        /** Numeric ID of the type */
        public int id;
        /** Name (system) of the type */
        public string name;

        private static List<RangeType> values = new List<RangeType>() {
            NUMBER,
            MIN,
            MAX,
            PERCENT,
            PERCENTILE,
            UNALLOCATED,
            FORMULA,
            AUTOMIN,
            AUTOMAX
        };
        public static List<RangeType> Values()
        {
            return values;
        }
        public override string ToString()
        {
            return id + " - " + name;
        }
        public override bool Equals(object obj)
        {
            if (obj == null || obj is not RangeType other)
            {
                return false;
            }

            return this.id == other.id && this.name == other.name;
        }

        public override int GetHashCode()
        {
            return id.GetHashCode() ^ name.GetHashCode();
        }

        public static RangeType ById(int id)
        {
            return Values()[id - 1]; // 1-based IDs
        }
        public static RangeType ByName(string name)
        {
            foreach (RangeType t in Values())
            {
                if (t.name.Equals(name)) return t;
            }
            return null;
        }

        private RangeType(int id, string name)
        {
            this.id = id; this.name = name;
        }
    }
}