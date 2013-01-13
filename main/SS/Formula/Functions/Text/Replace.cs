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
 * Created on May 15, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;
    using System.Text;
    using NPOI.SS.Formula.Eval;

    /**
     * An implementation of the Replace function:
     * Replaces part of a text string based on the number of Chars 
     * you specify, with another text string.
     * @author Manda Wilson &lt; wilson at c bio dot msk cc dot org &gt;
     */
    public class Replace : TextFunction
    {

        /**
         * Replaces part of a text string based on the number of Chars 
         * you specify, with another text string.
         * 
         * @see org.apache.poi.hssf.record.formula.eval.Eval
         */
        public override ValueEval EvaluateFunc(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            if (args.Length != 4)
            {
                return ErrorEval.VALUE_INVALID;
            }

            String oldStr = EvaluateStringArg(args[0], srcCellRow, srcCellCol);
            int startNum = EvaluateIntArg(args[1], srcCellRow, srcCellCol);
            int numChars = EvaluateIntArg(args[2], srcCellRow, srcCellCol);
            String newStr = EvaluateStringArg(args[3], srcCellRow, srcCellCol);

            if (startNum < 1 || numChars < 0)
            {
                return ErrorEval.VALUE_INVALID;
            }
            StringBuilder strBuff = new StringBuilder(oldStr);
            // remove any characters that should be replaced
            if (startNum <= oldStr.Length && numChars != 0)
            {
                strBuff.Remove(startNum - 1, Math.Min(numChars, oldStr.Length - startNum + 1));
            }
            // now insert (or append) newStr
            if (startNum > strBuff.Length)
            {
                strBuff.Append(newStr);
            }
            else
            {
                strBuff.Insert(startNum - 1, newStr);
            }
            return new StringEval(strBuff.ToString());
        }

    }
}