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

    /**
      * Compute the interest portion of a payment.
      * 
      * @author Mike Argyriou micharg@gmail.com
      */
    public class PPMT : NumericFunction
    {


        protected override double Eval(ValueEval[] args, int srcCellRow, int srcCellCol)
        {

            if (args.Length < 4)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);

            double result;
            ValueEval v5 =null;
            ValueEval v6 =null;

            ValueEval v1 = OperandResolver.GetSingleValue(args[0], srcCellRow, srcCellCol);
            ValueEval v2 = OperandResolver.GetSingleValue(args[1], srcCellRow, srcCellCol);
            ValueEval v3 = OperandResolver.GetSingleValue(args[2], srcCellRow, srcCellCol);
            ValueEval v4 = OperandResolver.GetSingleValue(args[3], srcCellRow, srcCellCol);

            double interestRate = OperandResolver.CoerceValueToDouble(v1);
            int period = OperandResolver.CoerceValueToInt(v2);
            int numberPayments = OperandResolver.CoerceValueToInt(v3);
            double PV = OperandResolver.CoerceValueToDouble(v4);

            if(args.Length==4)
            {
                result = Finance.PPMT(interestRate, period, numberPayments, PV);
            }
            else if(args.Length==5)
            {
                v5=OperandResolver.GetSingleValue(args[4], srcCellRow, srcCellCol);
                double FV = OperandResolver.CoerceValueToDouble(v5);
                result = Finance.PPMT(interestRate, period, numberPayments, PV, FV);
            }
            else
            {
                v5=OperandResolver.GetSingleValue(args[4], srcCellRow, srcCellCol);
                v6 = OperandResolver.GetSingleValue(args[5], srcCellRow, srcCellCol);
                double FV = OperandResolver.CoerceValueToDouble(v5);
                int type = OperandResolver.CoerceValueToInt(v6);
                result = Finance.PPMT(interestRate, period, numberPayments, PV, FV,type);
            }

            CheckValue(result);

            return result;
        }



    }

}