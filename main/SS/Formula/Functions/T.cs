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
    using NPOI.SS.Formula.Eval;

    public class T : Fixed1ArgFunction
    {
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            ValueEval arg = arg0;
            if (arg is RefEval)
            {
                // always use the first sheet
                RefEval re = (RefEval)arg;
                arg = re.GetInnerValueEval(re.FirstSheetIndex);
            }
            else if (arg is AreaEval)
            {
                // when the arg is an area, choose the top left cell
                arg = ((AreaEval)arg).GetRelativeValue(0, 0);
            }

            if (arg is StringEval)
            {
                // Text values are returned unmodified
                return arg;
            }

            if (arg is ErrorEval)
            {
                // Error values also returned unmodified
                return arg;
            }
            // for all other argument types the result is empty string
            return StringEval.EMPTY_INSTANCE;
        }
    }
}