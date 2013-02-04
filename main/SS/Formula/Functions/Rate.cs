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
using System;
using NPOI.SS.Formula.Eval;
using System.Diagnostics;

namespace NPOI.SS.Formula.Functions
{
    /**
     * Implements the Excel Rate function
     */
    public class Rate : Function
    {
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length < 3)
            { //First 3 parameters are mandatory
                return ErrorEval.VALUE_INVALID;
            }

            double periods, payment, present_val, future_val = 0, type = 0, estimate = 0.1, rate;

            try
            {
                ValueEval v1 = OperandResolver.GetSingleValue(args[0], srcRowIndex, srcColumnIndex);
                ValueEval v2 = OperandResolver.GetSingleValue(args[1], srcRowIndex, srcColumnIndex);
                ValueEval v3 = OperandResolver.GetSingleValue(args[2], srcRowIndex, srcColumnIndex);
                ValueEval v4 = null;
                if (args.Length >= 4)
                    v4 = OperandResolver.GetSingleValue(args[3], srcRowIndex, srcColumnIndex);
                ValueEval v5 = null;
                if (args.Length >= 5)
                    v5 = OperandResolver.GetSingleValue(args[4], srcRowIndex, srcColumnIndex);
                ValueEval v6 = null;
                if (args.Length >= 6)
                    v6 = OperandResolver.GetSingleValue(args[5], srcRowIndex, srcColumnIndex);

                periods = OperandResolver.CoerceValueToDouble(v1);
                payment = OperandResolver.CoerceValueToDouble(v2);
                present_val = OperandResolver.CoerceValueToDouble(v3);
                if (args.Length >= 4)
                    future_val = OperandResolver.CoerceValueToDouble(v4);
                if (args.Length >= 5)
                    type = OperandResolver.CoerceValueToDouble(v5);
                if (args.Length >= 6)
                    estimate = OperandResolver.CoerceValueToDouble(v6);
                rate = CalculateRate(periods, payment, present_val, future_val, type, estimate);

                CheckValue(rate);
            }
            catch (EvaluationException e)
            {
                Debug.WriteLine(e.Message);
                Debug.WriteLine(e.StackTrace);
                return e.GetErrorEval();
            }

            return new NumberEval(rate);
        }

        private double CalculateRate(double nper, double pmt, double pv, double fv, double type, double guess)
        {
            //FROM MS http://office.microsoft.com/en-us/excel-help/rate-HP005209232.aspx
            int FINANCIAL_MAX_ITERATIONS = 20;//Bet accuracy with 128
            double FINANCIAL_PRECISION = 0.0000001;//1.0e-8

            double y, y0, y1, x0, x1 = 0, f = 0, i = 0;
            double rate = guess;
            if (Math.Abs(rate) < FINANCIAL_PRECISION)
            {
                y = pv * (1 + nper * rate) + pmt * (1 + rate * type) * nper + fv;
            }
            else
            {
                f = Math.Exp(nper * Math.Log(1 + rate));
                y = pv * f + pmt * (1 / rate + type) * (f - 1) + fv;
            }
            y0 = pv + pmt * nper + fv;
            y1 = pv * f + pmt * (1 / rate + type) * (f - 1) + fv;

            // find root by Newton secant method
            i = x0 = 0.0;
            x1 = rate;
            while ((Math.Abs(y0 - y1) > FINANCIAL_PRECISION) && (i < FINANCIAL_MAX_ITERATIONS))
            {
                rate = (y1 * x0 - y0 * x1) / (y1 - y0);
                x0 = x1;
                x1 = rate;

                if (Math.Abs(rate) < FINANCIAL_PRECISION)
                {
                    y = pv * (1 + nper * rate) + pmt * (1 + rate * type) * nper + fv;
                }
                else
                {
                    f = Math.Exp(nper * Math.Log(1 + rate));
                    y = pv * f + pmt * (1 / rate + type) * (f - 1) + fv;
                }

                y0 = y1;
                y1 = y;
                ++i;
            }
            return rate;
        }

        /**
         * Excel does not support infinities and NaNs, rather, it gives a #NUM! error in these cases
         * 
         * @throws EvaluationException (#NUM!) if result is NaN or Infinity
         */
        static void CheckValue(double result)
        {
            if (Double.IsNaN(result) || Double.IsInfinity(result))
            {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
        }
    }
}
