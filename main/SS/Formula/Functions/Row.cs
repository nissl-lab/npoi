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

    public class Row : Function0Arg, Function1Arg
    {

        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex)
        {
            return new NumberEval(srcRowIndex + 1);
        }
        public ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            int rnum;

            if (arg0 is AreaEval)
            {
                rnum = ((AreaEval)arg0).FirstRow;
            }
            else if (arg0 is RefEval)
            {
                rnum = ((RefEval)arg0).Row;
            }
            else
            {
                // anything else is not valid argument
                return ErrorEval.VALUE_INVALID;
            }

            return new NumberEval(rnum + 1);
        }
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            switch (args.Length)
            {
                case 1:
                    return Evaluate(srcRowIndex, srcColumnIndex, args[0]);
                case 0:
                    return new NumberEval(srcRowIndex + 1);
            }
            return ErrorEval.VALUE_INVALID;
        }

    }
}