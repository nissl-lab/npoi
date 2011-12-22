/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 8, 2005
 *
 */
namespace NPOI.SS.Formula.Eval
{

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public interface OperationEval : Eval
    {

        /*
         * Read this, this will make your work easier when coding 
         * an "Evaluate()"
         * 
         * Things to note when implementing Evaluate():
         * 1. Check the Length of operands
         *    (use "switch(operands[x])" if possible)
         * 
         * 2. The possible Evals that you can Get as args to Evaluate are one of:
         * NumericValueEval, StringValueEval, RefEval, AreaEval
         * 3. If it is RefEval, the innerValueEval could be one of:
         * NumericValueEval, StringValueEval, BlankEval
         * 4. If it is AreaEval, each of the values could be one of:
         * NumericValueEval, StringValueEval, BlankEval, RefEval
         * 
         * 5. For numeric functions/operations, keep the result in double
         * till the end and before returning a new NumberEval, Check to see
         * if the double is a NaN - if NaN, return ErrorEval.ERROR_503 
         */
        Eval Evaluate(Eval[] evals, int srcCellRow, short srcCellCol);

        int NumberOfOperands { get; }

        //int Type { get; }
    }
}