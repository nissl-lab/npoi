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

namespace NPOI.SS.Formula.Atp
{
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.Eval;
    using System;
    using NPOI.SS.UserModel;


    /**
     * Implementation of Excel 'Analysis ToolPak' function EDATE()<br/>
     *
     * Adds a specified number of months to the specified date.<p/>
     *
     * <b>Syntax</b><br/>
     * <b>EDATE</b>(<b>date</b>, <b>number</b>)
     *
     * <p/>
     *
     * @author Tomas Herceg
     */
    public class EDate : FreeRefFunction
    {

        public static FreeRefFunction Instance = new EDate();

        private EDate()
        {
            // enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            double result;

            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                double startDateAsNumber = GetValue(args[0]);
                NumberEval offsetInYearsValue = (NumberEval)args[1];
                int offsetInMonthAsNumber = (int)offsetInYearsValue.NumberValue;
                // resolve the arguments
                DateTime startDate = DateUtil.GetJavaDate(startDateAsNumber);

                DateTime resultDate = startDate.AddMonths(offsetInMonthAsNumber);
                result = DateUtil.GetExcelDate(resultDate);
                    
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
        private double GetValue(ValueEval arg)
        {
            if (arg is NumberEval)
            {
                return ((NumberEval)arg).NumberValue;
            }
            if (arg is RefEval)
            {
                ValueEval innerValueEval = ((RefEval)arg).InnerValueEval;
                return ((NumberEval)innerValueEval).NumberValue;
            }
            throw new EvaluationException(ErrorEval.REF_INVALID);
        }
    }
}