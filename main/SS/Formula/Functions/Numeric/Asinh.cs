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
     * Support for hyperbolic trig functions was Added as a part of
     * Java distribution only in JDK1.5. This class uses custom
     * naive implementation based on formulas at:
     * http://www.math2.org/math/trig/hyperbolics.htm
     * These formulas seem to agree with excel's implementation.
     *
     */
    public class Asinh : OneArg
    {
        public override double Evaluate(double d)
        {
            return MathX.Asinh(d);
        }

    }
}