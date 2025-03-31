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
    using System.Text; 
using Cysharp.Text;
    using NPOI.SS.Formula.Eval;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    public class Concatenate : TextFunction
    {


        public override ValueEval EvaluateFunc(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            using var sb = ZString.CreateStringBuilder();

            for(int i=0, iSize=args.Length; i<iSize; i++) {
				sb.Append(EvaluateStringArg(args[i], srcCellRow, srcCellCol));
			}
			return new StringEval(sb.ToString());

        }
    }
}