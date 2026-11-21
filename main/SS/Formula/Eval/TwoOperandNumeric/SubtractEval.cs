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
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public class SubtractEval : TwoOperandNumericOperation
    {

        public override double Evaluate(double d0, double d1)
        {
            long bits = System.BitConverter.DoubleToInt64Bits(d0);
            bool negativeZero = bits == unchecked((long)0x8000000000000000L);
            decimal dec0 = (decimal)d0;
            decimal dec1 = (decimal)d1;
            double result = decimal.ToDouble(dec0 - dec1);
            if (result == 0.0 && negativeZero)
            {
                return -0.0;
            }
            return result;
        }
    }
}