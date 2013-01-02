/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
namespace NPOI.SS.Format
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;




    /**
     * This object represents a condition in a cell format.
     *
     * @author Ken Arnold, Industrious Media LLC
     */
    public abstract class CellFormatCondition
    {
        private const int LT = 0;
        private const int LE = 1;
        private const int GT = 2;
        private const int GE = 3;
        private const int EQ = 4;
        private const int NE = 5;

        private static Dictionary<String, int> TESTS;

        static CellFormatCondition()
        {
            TESTS = new Dictionary<String, int>();
            TESTS.Add("<", LT);
            TESTS.Add("<=", LE);
            TESTS.Add(">", GT);
            TESTS.Add(">=", GE);
            TESTS.Add("=", EQ);
            TESTS.Add("==", EQ);
            TESTS.Add("!=", NE);
            TESTS.Add("<>", NE);
        }
        private class LT_CellFormatCondition : CellFormatCondition
        {
            double _c;
            public LT_CellFormatCondition(double c)
            {
                _c = c;
            }
            public override bool Pass(double value)
            {
                return value < _c;
            }
        }
        private class LE_CellFormatCondition : CellFormatCondition
        {
            double _c;
            public LE_CellFormatCondition(double c)
            {
                _c = c;
            }
            public override bool Pass(double value)
            {
                return value <= _c;
            }
        }
        private class GT_CellFormatCondition : CellFormatCondition
        {
            double _c;
            public GT_CellFormatCondition(double c)
            {
                _c = c;
            }
            public override bool Pass(double value)
            {
                return value > _c;
            }
        }
        private class GE_CellFormatCondition : CellFormatCondition
        {
            double _c;
            public GE_CellFormatCondition(double c)
            {
                _c = c;
            }
            public override bool Pass(double value)
            {
                return value >= _c;
            }
        }
        private class EQ_CellFormatCondition : CellFormatCondition
        {
            double _c;
            public EQ_CellFormatCondition(double c)
            {
                _c = c;
            }
            public override bool Pass(double value)
            {
                return value == _c;
            }
        }
        private class NE_CellFormatCondition : CellFormatCondition
        {
            double _c;
            public NE_CellFormatCondition(double c)
            {
                _c = c;
            }
            public override bool Pass(double value)
            {
                return value != _c;
            }
        }
        /**
         * Returns an instance of a condition object.
         *
         * @param opString The operator as a string.  One of <tt>"&lt;"</tt>,
         *                 <tt>"&lt;="</tt>, <tt>">"</tt>, <tt>">="</tt>,
         *                 <tt>"="</tt>, <tt>"=="</tt>, <tt>"!="</tt>, or
         *                 <tt>"&lt;>"</tt>.
         * @param constStr The constant (such as <tt>"12"</tt>).
         *
         * @return A condition object for the given condition.
         */
        public static CellFormatCondition GetInstance(String opString,
                String constStr) {

            if (!TESTS.ContainsKey(opString))
                throw new ArgumentException("Unknown test: " + opString);
            int test = TESTS[(opString)];

            double c = Double.Parse(constStr, CultureInfo.InvariantCulture);

            switch (test)
            {
                case LT:
                    return new LT_CellFormatCondition(c);
                case LE:
                    return new LE_CellFormatCondition(c);
                case GT:
                    return new GT_CellFormatCondition(c);
                case GE:
                    return new GE_CellFormatCondition(c);
                case EQ:
                    return new EQ_CellFormatCondition(c);
                case NE:
                    return new NE_CellFormatCondition(c);
                default:
                    throw new ArgumentException(
                            "Cannot create for test number " + test + "(\"" + opString +
                                    "\")");
            }
        }

        /**
         * Returns <tt>true</tt> if the given value passes the constraint's test.
         *
         * @param value The value to compare against.
         *
         * @return <tt>true</tt> if the given value passes the constraint's test.
         */
        public abstract bool Pass(double value);
    }
}