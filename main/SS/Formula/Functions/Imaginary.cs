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
    using System.Text.RegularExpressions;
    /**
     * Implementation for Excel IMAGINARY() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>IMAGINARY  </b>(<b>Inumber</b>)<br/>
     * <p/>
     * Returns the imaginary coefficient of a complex number in x + yi or x + yj text format.
     * <p/>
     * Inumber     is a complex number for which you want the imaginary coefficient.
     * <p/>
     * Remarks
     * <ul>
     * <li>Use COMPLEX to convert real and imaginary coefficients into a complex number.</li>
     * </ul>
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Imaginary : Fixed1ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new Imaginary();

        public static String GROUP1_REAL_SIGN_REGEX = "([+-]?)";
        public static String GROUP2_REAL_INTEGER_OR_DOUBLE_REGEX = "([0-9]+\\.[0-9]+|[0-9]*)";
        public static String GROUP3_IMAGINARY_SIGN_REGEX = "([+-]?)";
        public static String GROUP4_IMAGINARY_INTEGER_OR_DOUBLE_REGEX = "([0-9]+\\.[0-9]+|[0-9]*)";
        public static String GROUP5_IMAGINARY_GROUP_REGEX = "([ij]?)";

        public static Regex COMPLEX_NUMBER_PATTERN
                = new Regex(GROUP1_REAL_SIGN_REGEX + GROUP2_REAL_INTEGER_OR_DOUBLE_REGEX +
                GROUP3_IMAGINARY_SIGN_REGEX + GROUP4_IMAGINARY_INTEGER_OR_DOUBLE_REGEX + GROUP5_IMAGINARY_GROUP_REGEX);

        public static int GROUP1_REAL_SIGN = 1;
        public static int GROUP2_IMAGINARY_INTEGER_OR_DOUBLE = 2;
        public static int GROUP3_IMAGINARY_SIGN = 3;
        public static int GROUP4_IMAGINARY_INTEGER_OR_DOUBLE = 4;

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

            //Matcher m = COMPLEX_NUMBER_PATTERN.matcher(iNumber);
            //bool result = m.matches();
            System.Text.RegularExpressions.Match m = COMPLEX_NUMBER_PATTERN.Match(iNumber);
            bool result = m.Success && m.Groups[0].Length>0;

            String imaginary = "";
            if (result == true)
            {
                String imaginaryGroup = m.Groups[5].Value;
                bool hasImaginaryPart = imaginaryGroup.Equals("i") || imaginaryGroup.Equals("j");

                if (imaginaryGroup.Length == 0)
                {
                    return new StringEval(Convert.ToString(0));
                }

                if (hasImaginaryPart)
                {
                    String sign = "";
                    String imaginarySign = m.Groups[(GROUP3_IMAGINARY_SIGN)].Value;
                    if (imaginarySign.Length != 0 && !(imaginarySign.Equals("+")))
                    {
                        sign = imaginarySign;
                    }

                    String groupImaginaryNumber = m.Groups[(GROUP4_IMAGINARY_INTEGER_OR_DOUBLE)].Value;
                    if (groupImaginaryNumber.Length != 0)
                    {
                        imaginary = sign + groupImaginaryNumber;
                    }
                    else
                    {
                        imaginary = sign + "1";
                    }
                }
            }
            else
            {
                return ErrorEval.NUM_ERROR;
            }

            return new StringEval(imaginary);
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