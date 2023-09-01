/*
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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.Formula.Functions
{
    public abstract class NumberListFuncBase : FreeRefFunction
    {
        public bool AllowEmptyList { get; set; } = false;
        public ValueEval ErrorOnEmptyList { get; set; } = ErrorEval.DIV_ZERO;
        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length == 0)
                return ErrorEval.VALUE_INVALID;

            try
            {
                var list = new List<double>();

                foreach (var arg in args)
                {
                    switch (arg)
                    {
                        case AreaEval ae:
                            ValueEvaluationHelper.GetArrayValues(ae, list);
                            break;
                        case NumericValueEval:
                        case RefEval:
                            var val = ValueEvaluationHelper.GetScalarValue(arg);
                            if (val.HasValue)
                                list.Add(val.Value);
                            break;
                        default:
                            return ErrorEval.VALUE_INVALID;
                    }
                }

                if (!AllowEmptyList && list.Count == 0)
                    return ErrorOnEmptyList;

                var result = CalculateFromNumberList(list);
                return result == 0.0 ? NumberEval.ZERO : new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }

        public abstract double CalculateFromNumberList(List<double> list);
    }
}
