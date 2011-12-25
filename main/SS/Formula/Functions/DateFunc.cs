/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    using System.Globalization;
    using NPOI.SS.UserModel;

    /**
     * @author Pavel Krupets (pkrupets at palmtreebusiness dot com)
     */
    public class DateFunc : Fixed3ArgFunction
    {
        public static Function instance = new DateFunc();

        private DateFunc()
        { 
            
        }
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,ValueEval arg2)
        {
            double result;
            try
            {
                double d0 = NumericFunction.SingleOperandEvaluate(arg0, srcRowIndex, srcColumnIndex);
                double d1 = NumericFunction.SingleOperandEvaluate(arg1, srcRowIndex, srcColumnIndex);
                double d2 = NumericFunction.SingleOperandEvaluate(arg2, srcRowIndex, srcColumnIndex);
                result = Evaluate(GetYear(d0), (int)d1, (int)d2);
                NumericFunction.CheckValue(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }
        public double Evaluate(int year, int month, int day)
        {
            if (year < 0 || month < 0 || day < 0)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }


            if (year == 1900 && month == 2 && day == 29)
            {
                return 60.0;
            }

            //see Microsoft KB214326
            //http://support.microsoft.com/kb/214326/en-us
            if (year == 1900)
            {
                if ((month == 1 && day >= 60) ||
                    (month == 2 && day >= 30))
                {
                    day--;
                }
            }

            return DateUtil.GetExcelDate(year, month, day, 0, 0, 0, false); // XXX fix 1900/1904 problem
        }

        private int GetYear(double d)
        {
            int year = (int)d;

            if (year < 0)
            {
                return -1;
            }

            return year < 1900 ? 1900 + year : year;
        }

    }
}