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
    using System.Text;
    /**
     * Implementation for Excel WeekNum() function.<p/>
     * <p/>
     * <b>Syntax</b>:<br/> <b>WeekNum  </b>(<b>Serial_num</b>,<b>Return_type</b>)<br/>
     * <p/>
     * Returns a number that indicates where the week falls numerically within a year.
     * <p/>
     * <p/>
     * Serial_num     is a date within the week. Dates should be entered by using the DATE function,
     * or as results of other formulas or functions. For example, use DATE(2008,5,23)
     * for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
     * Return_type     is a number that determines on which day the week begins. The default is 1.
     * 1	Week begins on Sunday. Weekdays are numbered 1 through 7.
     * 2	Week begins on Monday. Weekdays are numbered 1 through 7.
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Roman : Fixed2ArgFunction
    {

        //M (1000), CM (900), D (500), CD (400), C (100), XC (90), L (50), XL (40), X (10), IX (9), V (5), IV (4) and I (1).
        public static int[] VALUES = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        public static String[] ROMAN = new String[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval numberVE, ValueEval formVE)
        {
            int number = 0;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(numberVE, srcRowIndex, srcColumnIndex);
                number = OperandResolver.CoerceValueToInt(ve);
            }
            catch (EvaluationException)
            {
                return ErrorEval.VALUE_INVALID;
            }
            if (number < 0)
            {
                return ErrorEval.VALUE_INVALID;
            }
            if (number > 3999)
            {
                return ErrorEval.VALUE_INVALID;
            }
            if (number == 0)
            {
                return new StringEval("");
            }

            int form = 0;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(formVE, srcRowIndex, srcColumnIndex);
                form = OperandResolver.CoerceValueToInt(ve);
            }
            catch (EvaluationException)
            {
                return ErrorEval.NUM_ERROR;
            }

            if (form > 4 || form < 0)
            {
                return ErrorEval.VALUE_INVALID;
            }

            String result = this.integerToRoman(number);

            if (form == 0)
            {
                return new StringEval(result);
            }

            return new StringEval(MakeConcise(result, form));
        }

        /**
         * Classic conversion.
         *
         * @param number
         * @return
         */
        private String integerToRoman(int number)
        {
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < 13; i++)
            {
                while (number >= VALUES[i])
                {
                    number -= VALUES[i];
                    result.Append(ROMAN[i]);
                }
            }
            return result.ToString();
        }

        /**
         * Use conversion rule to factor some parts and make them more concise
         *
         * @param result
         * @param form
         * @return
         */
        public String MakeConcise(String result, int form)
        {
            if (form > 0)
            {
                result = result.Replace("XLV", "VL"); //45
                result = result.Replace("XCV", "VC"); //95
                result = result.Replace("CDL", "LD"); //450
                result = result.Replace("CML", "LM"); //950
                result = result.Replace("CMVC", "LMVL"); //995
            }
            if (form == 1)
            {
                result = result.Replace("CDXC", "LDXL"); //490
                result = result.Replace("CDVC", "LDVL"); //495
                result = result.Replace("CMXC", "LMXL"); //990
                result = result.Replace("XCIX", "VCIV"); //99
                result = result.Replace("XLIX", "VLIV"); //49
            }
            if (form > 1)
            {
                result = result.Replace("XLIX", "IL"); //49
                result = result.Replace("XCIX", "IC"); //99
                result = result.Replace("CDXC", "XD"); //490
                result = result.Replace("CDVC", "XDV"); //495
                result = result.Replace("CDIC", "XDIX"); //499
                result = result.Replace("LMVL", "XMV"); //995
                result = result.Replace("CMIC", "XMIX"); //999
                result = result.Replace("CMXC", "XM"); // 990
            }
            if (form > 2)
            {
                result = result.Replace("XDV", "VD");  //495
                result = result.Replace("XDIX", "VDIV"); //499
                result = result.Replace("XMV", "VM"); // 995
                result = result.Replace("XMIX", "VMIV"); //999
            }
            if (form == 4)
            {
                result = result.Replace("VDIV", "ID"); //499
                result = result.Replace("VMIV", "IM"); //999
            }

            return result;
        }
    }

}