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

using NPOI.SS.Formula.Eval;
using System;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using System.Globalization;
using NPOI.Util;
using System.Collections.Generic;
namespace NPOI.SS.Formula.Functions
{

    /**
     * Implementation for Excel FACTDOUBLE() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>FACTDOUBLE  </b>(<b>number</b>)<br/>
     * <p/>
     * Returns the double factorial of a number.
     * <p/>
     * Number is the value for which to return the double factorial. If number is not an integer, it is truncated.
     * <p/>
     * Remarks
     * <ul>
     * <li>If number is nonnumeric, FACTDOUBLE returns the #VALUE! error value.</li>
     * <li>If number is negative, FACTDOUBLE returns the #NUM! error value.</li>
     * </ul>
     * Use a cache for more speed of previously calculated factorial
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class FactDouble : Fixed1ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new FactDouble();

        //Caching of previously calculated factorial for speed
        static Dictionary<int, BigInteger> cache = new Dictionary<int, BigInteger>();

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval numberVE)
        {
            int number;
            try
            {
                number = OperandResolver.CoerceValueToInt(numberVE);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }

            if (number < 0)
            {
                return ErrorEval.NUM_ERROR;
            }

            return new NumberEval(factorial(number).LongValue());
        }

        public static BigInteger factorial(int n)
        {
            if (n == 0 || n < 0)
            {
                return BigInteger.One;
            }

            if (cache.ContainsKey(n))
            {
                return cache[(n)];
            }

            BigInteger result = BigInteger.ValueOf(n).Multiply(factorial(n - 2));
            cache.Add(n, result);
            return result;
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0]);
        }
    }
}