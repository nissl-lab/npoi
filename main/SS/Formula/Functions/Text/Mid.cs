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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;

    /// <summary>
    /// An implementation of the MID function
    /// MID returns a specific number of
    /// Chars from a text string, starting at the specified position.
    /// @author Manda Wilson &lt; wilson at c bio dot msk cc dot org;
    /// </summary>
    public class Mid : TextFunction
    {
        public override ValueEval EvaluateFunc(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            if (args.Length != 3)
            {
                return ErrorEval.VALUE_INVALID;
            }

            String text = EvaluateStringArg(args[0], srcCellRow, srcCellCol);
            int startCharNum = EvaluateIntArg(args[1], srcCellRow, srcCellCol);
            int numChars = EvaluateIntArg(args[2], srcCellRow, srcCellCol);
            int startIx = startCharNum - 1; // convert to zero-based

            // Note - for start_num arg, blank/zero causes error(#VALUE!),
            // but for num_chars causes empty string to be returned.
            if (startIx < 0)
            {
                return ErrorEval.VALUE_INVALID;
            }
            if (numChars < 0)
            {
                return ErrorEval.VALUE_INVALID;
            }
            int len = text.Length;
            if (numChars < 0 || startIx > len)
            {
                return new StringEval("");
            }
            int endIx = Math.Min(startIx + numChars, len);
            String result = text.Substring(startIx, endIx-startIx);
            return new StringEval(result);
        }
    }
}