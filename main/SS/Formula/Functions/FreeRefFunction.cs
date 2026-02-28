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

namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula;

    /**
     * For most Excel functions, involving references ((cell, area), (2d, 3d)), the references are 
     * passed in as arguments, and the exact location remains fixed.  However, a select few Excel
     * functions have the ability to access cells that were not part of any reference passed as an
     * argument.<br/>
     * Two important functions with this feature are <b>INDIRECT</b> and <b>OFFSet</b><p/>
     *  
     * In POI, the <c>HSSFFormulaEvaluator</c> Evaluates every cell in each reference argument before
     * calling the function.  This means that functions using fixed references do not need access to
     * the rest of the workbook to execute.  Hence the <c>Evaluate()</c> method on the common
     * interface <c>Function</c> does not take a workbook parameter.  
     * 
     * This interface recognises the requirement of some functions to freely Create and Evaluate 
     * references beyond those passed in as arguments.
     * 
     * @author Josh Micich
     */
    public interface FreeRefFunction
    {
        /**
         * @param args the pre-Evaluated arguments for this function. args is never <code>null</code>,
         *             nor are any of its elements.
         * @param ec primarily used to identify the source cell Containing the formula being Evaluated.
         *             may also be used to dynamically create reference evals.
         * @return never <code>null</code>. Possibly an instance of <c>ErrorEval</c> in the case of
         * a specified Excel error (Exceptions are never thrown to represent Excel errors).
         */
        ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec);
    }
}