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
    public class Even : OneArg
    {
        private const long PARITY_MASK = unchecked((long)0xFFFFFFFFFFFFFFFEL);

        public override double Evaluate(double d)
        {
            if (d == 0)
            {
                return 0;
            }
            long result;
            if (d > 0)
            {
                result = calcEven(d);
            }
            else
            {
                result = -calcEven(-d);
            }
            return result;
        }

        private static long calcEven(double d)
        {
            long x = ((long)d) & PARITY_MASK;
            if (x == d)
            {
                return x;
            }
            return x + 2;
        }

    }
}