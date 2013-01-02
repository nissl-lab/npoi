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

using System.Collections.Generic;
using NPOI.SS.Formula.Eval;
using System;
using NPOI.SS.UserModel;
namespace NPOI.SS.Formula.Atp
{
    /**
     * Evaluator for formula arguments.
     * 
     * @author jfaenomoto@gmail.com
     */
    internal class ArgumentsEvaluator
    {

        public static ArgumentsEvaluator instance = new ArgumentsEvaluator();

        private ArgumentsEvaluator()
        {
            // enforces singleton
        }

        /**
         * Evaluate a generic {@link ValueEval} argument to a double value that represents a date in POI.
         * 
         * @param arg {@link ValueEval} an argument.
         * @param srcCellRow number cell row.
         * @param srcCellCol number cell column.
         * @return a double representing a date in POI.
         * @throws EvaluationException exception upon argument evaluation.
         */
        public double EvaluateDateArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, srcCellRow, (short)srcCellCol);

            if (ve is StringEval)
            {
                String strVal = ((StringEval)ve).StringValue;
                Double dVal = OperandResolver.ParseDouble(strVal);
                if (!Double.IsNaN(dVal))
                {
                    return dVal;
                }
                DateTime dt = DateParser.ParseDate(strVal);
                return DateUtil.GetExcelDate(dt, false);
            }
            return OperandResolver.CoerceValueToDouble(ve);
        }

        /**
         * Evaluate a generic {@link ValueEval} argument to an array of double values that represents dates in POI.
         * 
         * @param arg {@link ValueEval} an argument.
         * @param srcCellRow number cell row.
         * @param srcCellCol number cell column.
         * @return an array of doubles representing dates in POI.
         * @throws EvaluationException exception upon argument evaluation.
         */
        public double[] EvaluateDatesArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            if (arg == null)
            {
                return new double[0];
            }

            if (arg is StringEval)
            {
                return new double[] { EvaluateDateArg(arg, srcCellRow, srcCellCol) };
            }
            else if (arg is AreaEvalBase)
            {
                List<Double> valuesList = new List<Double>();
                AreaEvalBase area = (AreaEvalBase)arg;
                for (int i = area.FirstRow; i <= area.LastRow; i++)
                {
                    for (int j = area.FirstColumn; j <= area.LastColumn; j++)
                    {
                        valuesList.Add(EvaluateDateArg(area.GetValue(i, j), i, j));
                    }
                }
                double[] values = new double[valuesList.Count];
                for (int i = 0; i < valuesList.Count; i++)
                {
                    values[i] = valuesList[(i)];
                }
                return values;
            }
            return new double[] { OperandResolver.CoerceValueToDouble(arg) };
        }

        /**
         * Evaluate a generic {@link ValueEval} argument to a double value.
         * 
         * @param arg {@link ValueEval} an argument.
         * @param srcCellRow number cell row.
         * @param srcCellCol number cell column.
         * @return a double value.
         * @throws EvaluationException exception upon argument evaluation.
         */
        public double EvaluateNumberArg(ValueEval arg, int srcCellRow, int srcCellCol)
        {
            if (arg == null)
            {
                return 0f;
            }

            return OperandResolver.CoerceValueToDouble(arg);
        }
    }

}