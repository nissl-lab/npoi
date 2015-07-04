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

namespace NPOI.SS.Formula.Functions
{
    public class Maxa : MinaMaxa
    {
        protected internal override double Evaluate(double[] values)
        {
            return values.Length > 0 ? MathX.Max(values) : 0;
        }
    }
    public class Mina : MinaMaxa
    {
         protected internal override double Evaluate(double[] values)
        {
            return values.Length > 0 ? MathX.Min(values) : 0;
        }
    }
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * 
     */
    public abstract class MinaMaxa : MultiOperandNumericFunction
    {

        protected MinaMaxa()
            : base(true, true)
        {

        }

        public static readonly Function MAXA = new Maxa();
        public static readonly Function MINA = new Mina();
    }
}
