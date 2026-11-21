/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.SS.Formula.Functions
{
    using ExtendedNumerics;
    using NPOI.SS.Formula.Eval;
    using System.Numerics;

    /// <summary>
    /// Implementation of the DAverage function:
    /// Gets the average value of a column in an area with given conditions.
    /// </summary>
    public sealed class DAverage : IDStarAlgorithm
    {
        private long count;
        private double total;
        public bool ProcessMatch(ValueEval eval)
        {
            if(eval is NumericValueEval valueEval)
            {
                count++;
                total += valueEval.NumberValue;
            }
            return true;
        }
        public ValueEval Result
        {
            get
            {
                return count == 0 ? NumberEval.ZERO : new NumberEval(GetAverage());
            }
        }

        private double GetAverage()
        {
            return Divide(total, count);
        }

        private static double Divide(double total, long count)
        {
            return (double) BigDecimal.Divide(new BigDecimal(total), new BigInteger(count));
        }

        public bool AllowEmptyMatchField { get; } = false;
    }
}


