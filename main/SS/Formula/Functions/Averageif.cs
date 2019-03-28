/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


namespace NPOI.SS.Formula.Functions {
    using NPOI.SS.Formula.Eval;

    /**
     * Implementation for the Excel function SUMIF<p>
     *
     * Syntax : <br/>
     *  AVERAGEIF ( <b>range</b>, <b>criteria</b>, avg_range ) <br/>
     *    <table border="0" cellpadding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><th>range</th><td>The range over which criteria is applied.  Also used for included values when the third parameter is not present</td></tr>
     *      <tr><th>criteria</th><td>The value or expression used to filter rows from <b>range</b></td></tr>
     *      <tr><th>avg_range</th><td>Locates the top-left corner of the corresponding range of addends - values to be included (after being selected by the criteria)</td></tr>
     *    </table><br/>
     * </p>
     * @author Josh Micich
     */
    public class AverageIf : FreeRefFunction {

        public static FreeRefFunction instance = new AverageIf();

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) {
            if (args.Length > 3 || args.Length < 2) {
                return ErrorEval.VALUE_INVALID;
            }
            return AverageIfs.instance.Evaluate(new [] {  GetSumRange(args), args[0], args[1]}, ec);
        }

        private ValueEval GetSumRange(ValueEval[] args){
            try {
                return args[2];
            }
            catch {
                return args[0];
            }
        }
    }

}
