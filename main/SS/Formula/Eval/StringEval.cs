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

namespace NPOI.SS.Formula.Eval
{
    using System;
    using System.Text;
    using NPOI.SS.Formula.PTG;


    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public class StringEval : StringValueEval
    {
        public static readonly StringEval EMPTY_INSTANCE = new StringEval("");

        private String value;

        public StringEval(Ptg ptg):this(((StringPtg)ptg).Value)
        {
            
        }

        public StringEval(String value)
        {
            if (value == null)
            {
                throw new ArgumentException("value must not be null");
            }
            this.value = value;
        }

        public String StringValue
        {
            get { return value; }
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(value);
            sb.Append("]");
            return sb.ToString();
        }
    }
}