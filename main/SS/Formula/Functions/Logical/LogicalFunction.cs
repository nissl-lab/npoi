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
 * Created on Nov 25, 2006
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    public abstract class LogicalFunction : Fixed1ArgFunction
    {
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0)
        {
            ValueEval ve;
            try
            {
                ve = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                //if (false)
                //{
                //    // Note - it is more usual to propagate error codes straight to the result like this:
                //    return e.GetErrorEval();
                //    // but logical functions behave a little differently
                //}
                // this will usually cause a 'FALSE' result except for ISNONTEXT()
                ve = e.GetErrorEval();
            }
            return BoolEval.ValueOf(Evaluate(ve));

        }
        /**
         * @param arg any {@link ValueEval}, potentially {@link BlankEval} or {@link ErrorEval}.
         */
        protected abstract bool Evaluate(ValueEval arg);

        public static readonly Function ISLOGICAL = new Islogical();
        public static readonly Function ISNONTEXT = new Isnontext();
        public static readonly Function ISNUMBER = new Isnumber();
        public static readonly Function ISTEXT = new Istext();
        public static readonly Function ISBLANK = new Isblank();
        public static readonly Function ISERROR = new Iserror();
        public static readonly Function ISNA = new Isna();
        public static readonly Function ISREF = new Isref();
        public static readonly Function ISERR = new Iserr();
    }
}