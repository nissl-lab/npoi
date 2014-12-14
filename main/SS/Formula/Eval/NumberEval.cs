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
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Util;
    using System.Globalization;
    using System.Text;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public class NumberEval : NumericValueEval, StringValueEval
    {

        public static readonly NumberEval ZERO = new NumberEval(0);

        private double _value;
        private String _stringValue;


        public NumberEval(Ptg ptg)
        {
            if (ptg is IntPtg)
            {
                this._value = ((IntPtg)ptg).Value;
            }
            else if (ptg is NumberPtg)
            {
                this._value = ((NumberPtg)ptg).Value;
            }
        }

        public NumberEval(double value)
        {
            this._value = value;
        }

        public double NumberValue
        {
            get { return _value; }
        }

        public String StringValue
        {
            get
            {// TODO: limit to 15 decimal places
                if (_stringValue == null)
                    //MakeString();
                    _stringValue = NumberToTextConverter.ToText(_value);
                return _stringValue;
            }
        }

        protected void MakeString()
        {
            if (!double.IsNaN(_value))
            {
                double lvalue = Math.Round(_value);
                if (lvalue == _value)
                {
                    _stringValue = lvalue.ToString(CultureInfo.CurrentCulture);
                }
                else
                {
                    _stringValue = _value.ToString(CultureInfo.CurrentCulture);
                }
            }
        }
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(GetType().Name).Append(" [");
            sb.Append(this.StringValue);
            sb.Append("]");
            return sb.ToString();
        }
    }
}