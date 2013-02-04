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

    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;

    /**
     * Common logic for rendering formulas.<br/>
     * 
     * For POI internal use only
     * 
     * @author Josh Micich
     */
    public class FormulaRenderer
    {

        /**
         * Static method To convert an array of {@link Ptg}s in RPN order
         * To a human readable string format in infix mode.
         * @param book  used for defined names and 3D references
         * @param ptgs  must not be <c>null</c>
         * @return a human readable String
         */
        public static String ToFormulaString(IFormulaRenderingWorkbook book, Ptg[] ptgs)
        {
            if (ptgs == null || ptgs.Length == 0)
            {
                throw new ArgumentException("ptgs must not be null");
            }
            Stack stack = new Stack();

            for (int i = 0; i < ptgs.Length; i++)
            {
                Ptg ptg = ptgs[i];
                // TODO - what about MemNoMemPtg?
                if (ptg is MemAreaPtg || ptg is MemFuncPtg || ptg is MemErrPtg)
                {
                    // marks the start of a list of area expressions which will be naturally combined
                    // by their trailing operators (e.g. UnionPtg)
                    // TODO - Put comment and throw exception in ToFormulaString() of these classes
                    continue;
                }
                if (ptg is ParenthesisPtg)
                {
                    String contents = (String)stack.Pop();
                    stack.Push("(" + contents + ")");
                    continue;
                }
                if (ptg is AttrPtg)
                {
                    AttrPtg attrPtg = ((AttrPtg)ptg);
                    if (attrPtg.IsOptimizedIf || attrPtg.IsOptimizedChoose || attrPtg.IsSkip)
                    {
                        continue;
                    }
                    if (attrPtg.IsSpace)
                    {
                        // POI currently doesn't render spaces in formulas
                        continue;
                        // but if it ever did, care must be taken:
                        // tAttrSpace comes *before* the operand it applies To, which may be consistent
                        // with how the formula text appears but is against the RPN ordering assumed here
                    }
                    if (attrPtg.IsSemiVolatile)
                    {
                        // similar To tAttrSpace - RPN is violated
                        continue;
                    }
                    if (attrPtg.IsSum)
                    {
                        String[] operands = GetOperands(stack, attrPtg.NumberOfOperands);
                        stack.Push(attrPtg.ToFormulaString(operands));
                        continue;
                    }
                    throw new Exception("Unexpected tAttr: " + attrPtg.ToString());
                }

                if (ptg is WorkbookDependentFormula)
                {
                    WorkbookDependentFormula optg = (WorkbookDependentFormula)ptg;
                    stack.Push(optg.ToFormulaString(book));
                    continue;
                }
                if (!(ptg is OperationPtg))
                {
                    stack.Push(ptg.ToFormulaString());
                    continue;
                }

                OperationPtg o = (OperationPtg)ptg;
                String[] operands1 = GetOperands(stack, o.NumberOfOperands);
                stack.Push(o.ToFormulaString(operands1));
            }
            if (stack.Count == 0)
            {
                // inspection of the code above reveals that every stack.pop() is followed by a
                // stack.push(). So this is either an internal error or impossible.
                throw new InvalidOperationException("Stack underflow");
            }
            String result = (String)stack.Pop();
            if (stack.Count != 0)
            {
                // Might be caused by some Tokens like AttrPtg and Mem*Ptg, which really shouldn't
                // Put anything on the stack
                throw new InvalidOperationException("too much stuff left on the stack");
            }
            return result;
        }

        private static String[] GetOperands(Stack stack, int nOperands)
        {
            String[] operands = new String[nOperands];

            for (int j = nOperands - 1; j >= 0; j--)
            { // reverse iteration because args were pushed in-order
                if (stack.Count == 0)
                {
                    String msg = "Too few arguments supplied to operation. Expected (" + nOperands
                         + ") operands but got (" + (nOperands - j - 1) + ")";
                    throw new InvalidOperationException(msg);
                }
                operands[j] = (String)stack.Pop();
            }
            return operands;
        }
    }
}