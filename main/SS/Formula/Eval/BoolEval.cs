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
    using System;
    using System.Text;
    using NPOI.SS.Formula.PTG;


    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public class BoolEval : NumericValueEval, StringValueEval
    {

        private bool value;

        public static readonly BoolEval FALSE = new BoolEval(false);
        public static readonly BoolEval TRUE = new BoolEval(true);

        /**
         * Convenience method for the following:<br/>
         * <c>(b ? BoolEval.TRUE : BoolEval.FALSE)</c>
         * @return a <c>BoolEval</c> instance representing <c>b</c>.
         */
        public static BoolEval ValueOf(bool b)
        {
            // TODO - Find / Replace all occurrences
            return b ? TRUE : FALSE;
        }

        public BoolEval(Ptg ptg)
        {
            this.value = ((BoolPtg)ptg).Value;
        }

        private BoolEval(bool value)
        {
            this.value = value;
        }

        public bool BooleanValue
        {
            get { return value; }
        }

        public double NumberValue
        {
            get{return value ? 1 : 0;}
        }

        public String StringValue
        {
            get { return value ? "TRUE" : "FALSE"; }
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(StringValue);
            sb.Append("]");
            return sb.ToString();
        }
    }
}