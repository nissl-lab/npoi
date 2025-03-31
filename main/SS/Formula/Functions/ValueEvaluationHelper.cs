﻿/*
 *  ====================================================================
 *    Licensed to the collaborators of the NPOI project under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The collaborators license this file to You under the Apache License, Version 2.0
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
using NPOI.SS.Formula.Eval;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 
using Cysharp.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    internal static class ValueEvaluationHelper
    {
        public static double? GetScalarValue(ValueEval arg)
        {

            ValueEval eval;
            if (arg is RefEval re)
            {
                if (re.NumberOfSheets > 1)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                eval = re.GetInnerValueEval(re.FirstSheetIndex);
            }
            else
            {
                eval = arg;
            }

            if (eval is AreaEval ae)
            {
                // an area ref can work as a scalar value if it is 1x1
                if (!ae.IsColumn || !ae.IsRow)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                eval = ae.GetRelativeValue(0, 0);
            }

            if (eval is null)
            {
                throw new ArgumentException("parameter (eval) may not be null");
            }

            return GetSingleValue(eval);
        }
        public static void GetArrayValues(AreaEval evalArg, in List<double> numList)
        {
            int height = evalArg.LastRow - evalArg.FirstRow + 1;
            int width = evalArg.LastColumn - evalArg.FirstColumn + 1; // TODO - junit

            for (int rrIx = 0; rrIx < height; rrIx++)
            {
                for (int rcIx = 0; rcIx < width; rcIx++)
                {
                    var val = GetSingleValue(evalArg.GetRelativeValue(rrIx, rcIx));
                    if (val.HasValue)
                        numList.Add(val.Value);
                }
            }
        }
        private static double? GetSingleValue(ValueEval ve) => ve switch
        {
            NumericValueEval nve => nve.NumberValue,
            null or BlankEval or StringEval => null,
            ErrorEval ev => throw new EvaluationException(ev),
            _ => throw new RuntimeException($"Unexpected value eval class ({ve.GetType().Name})")
        };
    }
}
