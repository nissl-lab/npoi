/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using System.Collections;
    using System.Reflection;

    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.PTG;

    /**
     * This class Creates <c>OperationEval</c> instances To help evaluate <c>OperationPtg</c>
     * formula Tokens.
     * 
     * @author Josh Micich
     */
    class OperationEvaluatorFactory
    {
        private static Type[] OPERATION_CONSTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg) };

        private static Hashtable _instancesByPtgClass = InitialiseInstancesMap();

        private OperationEvaluatorFactory()
        {
            // no instances of this class
        }

        private static Hashtable InitialiseInstancesMap()
        {
            Hashtable m = new Hashtable(32);
            Add(m, EqualPtg.instance, RelationalOperationEval.EqualEval);
            Add(m, GreaterEqualPtg.instance, RelationalOperationEval.GreaterEqualEval);
            Add(m, GreaterThanPtg.instance, RelationalOperationEval.GreaterThanEval);
            Add(m, LessEqualPtg.instance, RelationalOperationEval.LessEqualEval);
            Add(m, LessThanPtg.instance, RelationalOperationEval.LessThanEval);
            Add(m, NotEqualPtg.instance, RelationalOperationEval.NotEqualEval);

            Add(m, ConcatPtg.instance, ConcatEval.instance);
            Add(m, AddPtg.instance, TwoOperandNumericOperation.AddEval);
            Add(m, DividePtg.instance, TwoOperandNumericOperation.DivideEval);
            Add(m, MultiplyPtg.instance, TwoOperandNumericOperation.MultiplyEval);
            Add(m, PercentPtg.instance, PercentEval.instance);
            Add(m, PowerPtg.instance, TwoOperandNumericOperation.PowerEval);
            Add(m, SubtractPtg.instance, TwoOperandNumericOperation.SubtractEval);
            Add(m, UnaryMinusPtg.instance, UnaryMinusEval.instance);
            Add(m, UnaryPlusPtg.instance, UnaryPlusEval.instance);
            Add(m, RangePtg.instance, RangeEval.instance);
            Add(m, IntersectionPtg.instance, IntersectionEval.instance);
            return m;
        }

        private static void Add(Hashtable m, OperationPtg ptgKey,
            NPOI.SS.Formula.Functions.Function instance)
        {
            // make sure ptg has single private constructor because map lookups assume singleton keys
            ConstructorInfo[] cc = ptgKey.GetType().GetConstructors();
            if (cc.Length > 1 || (cc.Length > 0 && !cc[0].IsPrivate))
            {
                throw new Exception("Failed to verify instance ("
                        + ptgKey.GetType().Name + ") is a singleton.");
            }
            m[ptgKey] = instance;
        }


        /**
         * returns the OperationEval concrete impl instance corresponding
         * to the supplied operationPtg
         */
        public static ValueEval Evaluate(OperationPtg ptg, ValueEval[] args,
                OperationEvaluationContext ec)
        {
            if (ptg == null)
            {
                throw new ArgumentException("ptg must not be null");
            }
            NPOI.SS.Formula.Functions.Function result = _instancesByPtgClass[ptg] as NPOI.SS.Formula.Functions.Function;

            if (result != null)
            {
                return result.Evaluate(args, ec.RowIndex, (short)ec.ColumnIndex);
            }

            if (ptg is AbstractFunctionPtg)
            {
                AbstractFunctionPtg fptg = (AbstractFunctionPtg)ptg;
                int functionIndex = fptg.FunctionIndex;
                switch (functionIndex)
                {
                    case NPOI.SS.Formula.Function.FunctionMetadataRegistry.FUNCTION_INDEX_INDIRECT:
                        return Indirect.instance.Evaluate(args, ec);
                    case NPOI.SS.Formula.Function.FunctionMetadataRegistry.FUNCTION_INDEX_EXTERNAL:
                        return UserDefinedFunction.instance.Evaluate(args, ec);
                }

                return FunctionEval.GetBasicFunction(functionIndex).Evaluate(args, ec.RowIndex, ec.ColumnIndex);
            }
            throw new Exception("Unexpected operation ptg class (" + ptg.GetType().Name + ")");
        }
    }
}