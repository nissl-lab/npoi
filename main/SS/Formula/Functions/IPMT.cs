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
    using NPOI.SS.Formula.Eval;

    public class IPMT : NumericFunction
    {
        protected override double Eval(ValueEval[] args, int srcCellRow, int srcCellCol)
        {

            if (args.Length != 4)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            double result;

            ValueEval v1 = OperandResolver.GetSingleValue(args[0], srcCellRow, srcCellCol);
            ValueEval v2 = OperandResolver.GetSingleValue(args[1], srcCellRow, srcCellCol);
            ValueEval v3 = OperandResolver.GetSingleValue(args[2], srcCellRow, srcCellCol);
            ValueEval v4 = OperandResolver.GetSingleValue(args[3], srcCellRow, srcCellCol);

            double interestRate = OperandResolver.CoerceValueToDouble(v1);
            int period = OperandResolver.CoerceValueToInt(v2);
            int numberPayments = OperandResolver.CoerceValueToInt(v3);
            double PV = OperandResolver.CoerceValueToDouble(v4);

            result = Finance.IPMT(interestRate, period, numberPayments, PV);

            CheckValue(result);

            return result;
        }
    }
}