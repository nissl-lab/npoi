/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 6, 2005
 *
 */
using NPOI.SS.Formula.Eval;
using NPOI.SS.Util;
using System;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    public class Dollar : Function
    {
        protected static double singleOperandEvaluate(ValueEval arg, int srcRowIndex, int srcColumnIndex)
        {
            if (arg == null)
            {
                throw new ArgumentException("arg must not be null");
            }
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcRowIndex, srcColumnIndex);
            double result = OperandResolver.CoerceValueToDouble(ve);
            checkValue(result);
            return result;
        }
        public static void checkValue(double result)
        {
            if (Double.IsNaN(result) || Double.IsInfinity(result))
            {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
        }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length != 1 && args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }
            try
            {
                double val = singleOperandEvaluate(args[0], srcRowIndex, srcColumnIndex);
                double d1 = args.Length == 1 ? 2.0 : singleOperandEvaluate(args[1], srcRowIndex, srcColumnIndex);
                // second arg converts to int by truncating toward zero
                int nPlaces = (int)d1;

                if (nPlaces > 127)
                {
                    return ErrorEval.VALUE_INVALID;
                }
                if (nPlaces < 0)
                {
                    d1 = Math.Abs(d1);
                    double temp = val / Math.Pow(10, d1);
                    temp = Math.Round(temp, 0);
                    val = temp* Math.Pow(10, d1);
                }
                StringBuilder decimalPlacesFormat = new StringBuilder();
                if (nPlaces > 0)
                {
                    decimalPlacesFormat.Append('.');
                }
                for (int i = 0; i < nPlaces; i++)
                {
                    decimalPlacesFormat.Append('0');
                }
                StringBuilder decimalFormatString = new StringBuilder();
                decimalFormatString.Append("$#,##0").Append(decimalPlacesFormat)
                        .Append(";($#,##0").Append(decimalPlacesFormat).Append(')');

                DecimalFormat df = new DecimalFormat(decimalFormatString.ToString());

                return new StringEval(df.Format(val));
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }
}