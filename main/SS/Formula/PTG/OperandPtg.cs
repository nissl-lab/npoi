/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.PTG
{
    using System;
    using NPOI.Util;
    /**
     * @author Josh Micich
     */
    [Serializable]
    public abstract class OperandPtg : Ptg
    {

        /**
         * All Operand <c>Ptg</c>s are classifed ('relative', 'value', 'array')  
         */
        public override bool IsBaseToken
        {
            get { return false; }
        }
        public OperandPtg Copy()
        {
            try
            {
                return (OperandPtg)Clone();
            }
            catch (NotSupportedException e)
            {
                throw new RuntimeException(e);
            }
        }
    }
}

