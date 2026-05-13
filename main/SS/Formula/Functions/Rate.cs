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
using EFF = Excel.FinancialFunctions;
using NPOI.SS.Formula.Eval;
using NPOI.Util;

namespace NPOI.SS.Formula.Functions
{
    public class Rate : Function
    {
        private static readonly POILogger Logger = POILogFactory.GetLogger(typeof(Rate));

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            if (args.Length < 3)
                return ErrorEval.VALUE_INVALID;

            try
            {
                double nper = FinancialHelper.GetDoubleArg(args, 0, srcRowIndex, srcColumnIndex);
                double pmt = FinancialHelper.GetDoubleArg(args, 1, srcRowIndex, srcColumnIndex);
                double pv = FinancialHelper.GetDoubleArg(args, 2, srcRowIndex, srcColumnIndex);
                double fv = FinancialHelper.GetDoubleArg(args, 3, srcRowIndex, srcColumnIndex, 0.0);
                double type = FinancialHelper.GetDoubleArg(args, 4, srcRowIndex, srcColumnIndex, 0.0);
                double guess = FinancialHelper.GetDoubleArg(args, 5, srcRowIndex, srcColumnIndex, 0.1);

                double rate;
                try
                {
                    rate = EFF.Financial.Rate(nper, pmt, pv, fv, FinancialHelper.ToPaymentDue(type), guess);
                }
                catch
                {
                    return ErrorEval.NUM_ERROR;
                }

                CheckValue(rate);
                return new NumberEval(rate);
            }
            catch (EvaluationException e)
            {
                Logger.Log(POILogger.ERROR, "Can't evaluate rate function");
                return e.GetErrorEval();
            }
        }

        static void CheckValue(double result)
        {
            if (double.IsNaN(result) || double.IsInfinity(result))
                throw new EvaluationException(ErrorEval.NUM_ERROR);
        }
    }
}
