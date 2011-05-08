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
namespace NPOI.HSSF.Record.Formula.Eval
{
    using System;
    using NPOI.HSSF.Record.Formula;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public class NumberEval : NumericValueEval, StringValueEval
    {

        public static NumberEval ZERO = new NumberEval(0);

        private double value;
        private String stringValue;


        public NumberEval(Ptg ptg)
        {
            if (ptg is IntPtg)
            {
                this.value = ((IntPtg)ptg).Value;
            }
            else if (ptg is NumberPtg)
            {
                this.value = ((NumberPtg)ptg).Value;
            }
        }

        public NumberEval(double value)
        {
            this.value = value;
        }

        public double NumberValue
        {
            get { return value; }
        }

        public String StringValue
        {
            get
            {// TODO: limit to 15 decimal places
                if (stringValue == null)
                    MakeString();
                return stringValue;
            }
        }

        protected void MakeString()
        {
            if (!double.IsNaN(value))
            {
                double lvalue = Math.Round(value);
                if (lvalue == value)
                {
                    stringValue = lvalue.ToString();
                }
                else
                {
                    stringValue = value.ToString();
                }
            }
        }

    }
}