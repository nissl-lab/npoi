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
using NPOI.Util;

namespace NPOI.SS.Formula.Functions
{
    /**
     * Implements the Excel Rate function
     */
    public class Rate : Function
    {
        private static readonly POILogger Logger = POILogFactory.GetLogger(typeof(Rate));

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if(args.Length < 3)
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
                if(args.Length >= 4)
                    v4 = OperandResolver.GetSingleValue(args[3], srcRowIndex, srcColumnIndex);
                ValueEval v5 = null;
                if(args.Length >= 5)
                    v5 = OperandResolver.GetSingleValue(args[4], srcRowIndex, srcColumnIndex);
                ValueEval v6 = null;
                if(args.Length >= 6)
                    v6 = OperandResolver.GetSingleValue(args[5], srcRowIndex, srcColumnIndex);

                periods = OperandResolver.CoerceValueToDouble(v1);
                payment = OperandResolver.CoerceValueToDouble(v2);
                present_val = OperandResolver.CoerceValueToDouble(v3);
                if(args.Length >= 4)
                    future_val = OperandResolver.CoerceValueToDouble(v4);
                if(args.Length >= 5)
                    type = OperandResolver.CoerceValueToDouble(v5);
                if(args.Length >= 6)
                    estimate = OperandResolver.CoerceValueToDouble(v6);
                rate = CalculateRate(periods, payment, present_val, future_val, type, estimate);

                CheckValue(rate);
            }
            catch(EvaluationException e)
            {
                Logger.Log(POILogger.ERROR, "Can't evaluate rate function");
                return e.GetErrorEval();
            }

            return new NumberEval(rate);
        }

        private static double G_Div_GP(double r, double n, double p, double x, double y, double w)
        {
            double t1 = Math.Pow(r+1, n);
            double t2 = Math.Pow(r+1, n-1);
            return (y + t1*x + p*(t1 - 1)*(r*w + 1)/r) /
                    (n*t2*x - p*(t1 - 1)*(r*w + 1)/(Math.Pow(r, 2) + n*p*t2*(r*w + 1)/r +
                            p*(t1 - 1)*w/r));
        }

        /// <summary>
        /// Compute the rate of interest per period.
        /// </summary>
        /// <param name="nper">Number of compounding periods</param>
        /// <param name="pmt">Payment</param>
        /// <param name="pv">Present Value</param>
        /// <param name="fv">Future value</param>
        /// <param name="type">When payments are due ('begin' (1) or 'end' (0))</param>
        /// <param name="guess">Starting guess for solving the rate of interest</param>
        /// <returns>rate of interest per period or NaN if the solution didn't converge</returns>
        static double CalculateRate(double nper, double pmt, double pv, double fv, double type, double guess)
        {
            double tol = 1e-8;
            double maxiter = 100;

            double rn = guess;
            int iter = 0;
            bool close = false;

            while(iter < maxiter && !close)
            {
                double rnp1 = rn - G_Div_GP(rn, nper, pmt, pv, fv, type);
                double diff = Math.Abs(rnp1 - rn);
                close = diff < tol;
                iter += 1;
                rn = rnp1;

            }
            if(!close)
                return double.NaN;
            else
            {
                return rn;
            }
        }

        /**
         * Excel does not support infinities and NaNs, rather, it gives a #NUM! error in these cases
         * 
         * @throws EvaluationException (#NUM!) if result is NaN or Infinity
         */
        static void CheckValue(double result)
        {
            if (double.IsNaN(result) || double.IsInfinity(result))
            {
                throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
        }
    }
}
