/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Function;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Reflection;


    /**
     * This class Creates <c>OperationEval</c> instances to help evaluate <c>OperationPtg</c>
     * formula tokens.
     * 
     * @author Josh Micich
     */
    class OperationEvaluatorFactory
    {
        private static Dictionary<OperationPtg, Function> _instancesByPtgClass = InitialiseConstructorsMap();
        private OperationEvaluatorFactory()
        {
            // no instances of this class
        }

        private static Dictionary<OperationPtg, Function> InitialiseConstructorsMap()
        {
            Dictionary<OperationPtg, Function> m = new Dictionary<OperationPtg, Function>(32);

            m.Add(AddPtg.instance, TwoOperandNumericOperation.AddEval); // 0x03
            m.Add(SubtractPtg.instance, TwoOperandNumericOperation.SubtractEval); // 0x04
            m.Add(MultiplyPtg.instance, TwoOperandNumericOperation.MultiplyEval); // 0x05
            m.Add(DividePtg.instance, TwoOperandNumericOperation.DivideEval); // 0x06
            m.Add(PowerPtg.instance, TwoOperandNumericOperation.PowerEval); // 0x07
            m.Add(ConcatPtg.instance, ConcatEval.instance); // 0x08
            m.Add(LessThanPtg.instance, RelationalOperationEval.LessThanEval); // 0x09
            m.Add(LessEqualPtg.instance, RelationalOperationEval.LessEqualEval); // 0x0a
            m.Add(EqualPtg.instance, RelationalOperationEval.EqualEval); // 0x0b
            m.Add(GreaterEqualPtg.instance, RelationalOperationEval.GreaterEqualEval); // 0x0c
            m.Add(GreaterThanPtg.instance, RelationalOperationEval.GreaterThanEval); // 0x0D
            m.Add(NotEqualPtg.instance, RelationalOperationEval.NotEqualEval); // 0x0e
            m.Add(IntersectionPtg.instance, IntersectionEval.instance); // 0x0f
            m.Add(RangePtg.instance, RangeEval.instance); // 0x11
            m.Add(UnaryPlusPtg.instance, UnaryPlusEval.instance); // 0x12
            m.Add(UnaryMinusPtg.instance, UnaryMinusEval.instance); // 0x13
            m.Add(PercentPtg.instance, PercentEval.instance); // 0x14

            return m;
        }


        /// <summary>
        /// returns the OperationEval concrete impl instance corresponding
        /// to the supplied operationPtg
        /// </summary>
        public static ValueEval evaluate(OperationPtg ptg, ValueEval[] args,
                OperationEvaluationContext ec)
        {
            if(ptg == null)
            {
                throw new ArgumentException("ptg must not be null");
            }
            Function result = _instancesByPtgClass[ptg];

            if(result != null)
            {
                IEvaluationSheet evalSheet = ec.GetWorkbook().GetSheet(ec.SheetIndex);
                IEvaluationCell evalCell = evalSheet.GetCell(ec.RowIndex, ec.ColumnIndex);

                if(evalCell.IsPartOfArrayFormulaGroup && result is IArrayFunction)
                    return ((IArrayFunction) result).EvaluateArray(args, ec.RowIndex, ec.ColumnIndex);

                return result.Evaluate(args, ec.RowIndex, (short) ec.ColumnIndex);
            }

            if(ptg is AbstractFunctionPtg)
            {
                AbstractFunctionPtg fptg = (AbstractFunctionPtg)ptg;
                int functionIndex = fptg.FunctionIndex;
                switch(functionIndex)
                {
                    case FunctionMetadataRegistry.FUNCTION_INDEX_INDIRECT:
                        return Indirect.instance.Evaluate(args, ec);
                    case FunctionMetadataRegistry.FUNCTION_INDEX_EXTERNAL:
                        return UserDefinedFunction.instance.Evaluate(args, ec);
                }

                return FunctionEval.GetBasicFunction(functionIndex).Evaluate(args, ec.RowIndex, (short) ec.ColumnIndex);
            }
            throw new RuntimeException("Unexpected operation ptg class (" + ptg.GetType().Name + ")");
        }
    }
}