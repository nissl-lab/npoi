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

namespace NPOI.XSSF.UserModel
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;

    /**
     * Internal POI use only - parent of XSSF and SXSSF formula Evaluators
     */
    public abstract class BaseXSSFFormulaEvaluator : BaseFormulaEvaluator
    {

        protected BaseXSSFFormulaEvaluator(WorkbookEvaluator bookEvaluator)
            : base(bookEvaluator)
        {
        }
        protected override IRichTextString CreateRichTextString(String str)
        {
            return new XSSFRichTextString(str);
        }
        public override void NotifySetFormula(ICell cell)
        {
            _bookEvaluator.NotifyUpdateCell(new XSSFEvaluationCell((XSSFCell)cell));
        }
        public override void NotifyDeleteCell(ICell cell)
        {
            _bookEvaluator.NotifyDeleteCell(new XSSFEvaluationCell((XSSFCell)cell));
        }
        public override void NotifyUpdateCell(ICell cell)
        {
            _bookEvaluator.NotifyUpdateCell(new XSSFEvaluationCell((XSSFCell)cell));
        }

        /**
         * Turns a XSSFCell / SXSSFCell into a XSSFEvaluationCell
         */
        protected abstract IEvaluationCell ToEvaluationCell(ICell cell);

        /**
         * Returns a CellValue wrapper around the supplied ValueEval instance.
         */
        protected override CellValue EvaluateFormulaCellValue(ICell cell)
        {
            IEvaluationCell evalCell = ToEvaluationCell(cell);
            ValueEval eval = _bookEvaluator.Evaluate(evalCell);
            if (eval is NumberEval numberEval)
            {
                return new CellValue(numberEval.NumberValue);
            }
            if (eval is BoolEval be)
            {
                return CellValue.ValueOf(be.BooleanValue);
            }
            if (eval is StringEval ne)
            {
                return new CellValue(ne.StringValue);
            }
            if (eval is ErrorEval errorEval)
            {
                return CellValue.GetError(errorEval.ErrorCode);
            }
            throw new Exception("Unexpected eval class (" + eval.GetType().Name + ")");
        }
    }

}