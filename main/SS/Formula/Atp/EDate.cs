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
    class EDate : FreeRefFunction
    {

        public static FreeRefFunction Instance = new EDate();

        private EDate()
        {
            // enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            DateTime date;
            double numberOfMonths, result;

            if (args.Length != 2)
            {
                return ErrorEval.VALUE_INVALID;
            }

            try
            {
                // resolve the arguments
                date = DateUtil.GetJavaDate(OperandResolver.CoerceValueToDouble(OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec.ColumnIndex)));
                numberOfMonths = OperandResolver.CoerceValueToDouble(OperandResolver.GetSingleValue(args[1], ec.RowIndex, ec.ColumnIndex));

                // calculate the result date (Excel rounds the second argument always to zero; but we have be careful about negative numbers)
                DateTime resultDate = date.AddMonths((int)Math.Floor(Math.Abs(numberOfMonths)) * Math.Sign(numberOfMonths));
                result = DateUtil.GetExcelDate(resultDate);
                    
                NumericFunction.CheckValue(result);
                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }
}