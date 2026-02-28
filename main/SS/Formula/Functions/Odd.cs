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
 * Created on May 6, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public class Odd : OneArg
    {

        private const long PARITY_MASK = unchecked((long)0xFFFFFFFFFFFFFFFEL);

        public override double Evaluate(double d)
        {
            if (d == 0)
            {
                return 1;
            }
            long result;
            if (d > 0)
            {
                result = CalcOdd(d);
            }
            else
            {
                result = -CalcOdd(-d);
            }
            return result;
        }

        private static long CalcOdd(double d)
        {
            double dpm1 = d + 1;
            long x = ((long)dpm1) & PARITY_MASK;
            if (x == dpm1)
            {
                return x - 1;
            }
            return x + 1;
        }

    }
}