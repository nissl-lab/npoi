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
     * Implementation for Excel Bin2Dec() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>Bin2Dec  </b>(<b>number</b>)<br/>
     * <p/>
     * Converts a binary number to decimal.
     * <p/>
     * Number is the binary number you want to convert. Number cannot contain more than 10 characters (10 bits).
     * The most significant bit of number is the sign bit. The remaining 9 bits are magnitude bits.
     * Negative numbers are represented using two's-complement notation.
     * <p/>
     * Remark
     * If number is not a valid binary number, or if number contains more than 10 characters (10 bits),
     * BIN2DEC returns the #NUM! error value.
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Bin2Dec : Fixed1ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new Bin2Dec();

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval numberVE)
        {
            String number;
            if (numberVE is RefEval)
            {
                RefEval re = (RefEval)numberVE;
                number = OperandResolver.CoerceValueToString(re.GetInnerValueEval(re.FirstSheetIndex));
            }
            else
            {
                number = OperandResolver.CoerceValueToString(numberVE);
            }
            if (number.Length > 10)
            {
                return ErrorEval.NUM_ERROR;
            }

            String unsigned;

            //If the leftmost bit is 0 -- number is positive.
            bool isPositive;
            if (number.Length < 10)
            {
                unsigned = number;
                isPositive = true;
            }
            else
            {
                unsigned = number.Substring(1);
                isPositive = number.StartsWith("0");
            }

            String value;
            try
            {
                if (isPositive)
                {
                    //bit9*2^8 + bit8*2^7 + bit7*2^6 + bit6*2^5 + bit5*2^4+ bit3*2^2+ bit2*2^1+ bit1*2^0
                    int sum = getDecimalValue(unsigned);
                    value = sum.ToString();
                }
                else
                {
                    //The leftmost bit is 1 -- this is negative number
                    //Inverse bits [1-9]
                    String inverted = toggleBits(unsigned);
                    // Calculate decimal number
                    int sum = getDecimalValue(inverted);

                    //Add 1 to obtained number
                    sum++;

                    value = "-" + sum.ToString();
                }
            }
            catch (FormatException)
            {
                return ErrorEval.NUM_ERROR;
            }
            return new NumberEval(long.Parse(value));
        }

        private int getDecimalValue(String unsigned)
        {
            int sum = 0;
            int numBits = unsigned.Length;
            int power = numBits - 1;

            for (int i = 0; i < numBits; i++)
            {
                int bit = int.Parse(unsigned.Substring(i, 1));
                int term = (int)(bit * Math.Pow(2, power));
                sum += term;
                power--;
            }
            return sum;
        }

        private static String toggleBits(String s)
        {
            long i = Convert.ToInt64(s, 2);
            long i2 = i ^ ((1L << s.Length) - 1);
            String s2 = Convert.ToString(i2, 2);
            while (s2.Length < s.Length) s2 = '0' + s2;
            return s2;
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