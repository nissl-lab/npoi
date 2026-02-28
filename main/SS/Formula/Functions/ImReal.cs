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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;

    using NPOI.SS.Formula;

    /**
     * Implementation for Excel ImReal() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>ImReal  </b>(<b>Inumber</b>)<br/>
     * <p/>
     * Returns the real coefficient of a complex number in x + yi or x + yj text format.
     * <p/>
     * Inumber     A complex number for which you want the real coefficient.
     * <p/>
     * Remarks
     * <ul>
     * <li>If inumber is not in the form x + yi or x + yj, this function returns the #NUM! error value.</li>
     * <li>Use COMPLEX to convert real and imaginary coefficients into a complex number.</li>
     * </ul>
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class ImReal : Fixed1ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new ImReal();

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval inumberVE)
        {
            ValueEval veText1;
            try
            {
                veText1 = OperandResolver.GetSingleValue(inumberVE, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            String iNumber = OperandResolver.CoerceValueToString(veText1);

            System.Text.RegularExpressions.Match m = Imaginary.COMPLEX_NUMBER_PATTERN.Match(iNumber);
            //bool result = m.matches();
            bool result = m.Success && !string.IsNullOrEmpty(m.Groups[0].Value);

            String real = "";
            if (result == true)
            {
                String realGroup = m.Groups[(2)].Value;
                bool hasRealPart = realGroup.Length != 0;

                if (realGroup.Length == 0)
                {
                    return new StringEval(Convert.ToString(0));
                }

                if (hasRealPart)
                {
                    String sign = "";
                    String realSign = m.Groups[(Imaginary.GROUP1_REAL_SIGN)].Value;
                    if (realSign.Length != 0 && !(realSign.Equals("+")))
                    {
                        sign = realSign;
                    }

                    String groupRealNumber = m.Groups[(Imaginary.GROUP2_IMAGINARY_INTEGER_OR_DOUBLE)].Value;
                    if (groupRealNumber.Length != 0)
                    {
                        real = sign + groupRealNumber;
                    }
                    else
                    {
                        real = sign + "1";
                    }
                }
            }
            else
            {
                return ErrorEval.NUM_ERROR;
            }

            return new StringEval(real);
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