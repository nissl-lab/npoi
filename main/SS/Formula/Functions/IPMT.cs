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

namespace NPOI.SS.Formula.Functions
{
    using EFF = Excel.FinancialFunctions;
    using NPOI.SS.Formula.Eval;

    public class IPMT : NumericFunction
    {
        protected override double Eval(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            if (args.Length < 4 || args.Length > 6)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            double rate = FinancialHelper.GetDoubleArg(args, 0, srcCellRow, srcCellCol);
            double per = FinancialHelper.GetDoubleArg(args, 1, srcCellRow, srcCellCol);
            double nper = FinancialHelper.GetDoubleArg(args, 2, srcCellRow, srcCellCol);
            double pv = FinancialHelper.GetDoubleArg(args, 3, srcCellRow, srcCellCol);
            double fv = FinancialHelper.GetDoubleArg(args, 4, srcCellRow, srcCellCol, 0.0);
            double type = FinancialHelper.GetDoubleArg(args, 5, srcCellRow, srcCellCol, 0.0);

            double result = EFF.Financial.IPmt(rate, per, nper, pv, fv, FinancialHelper.ToPaymentDue(type));
            CheckValue(result);
            return result;
        }
    }
}