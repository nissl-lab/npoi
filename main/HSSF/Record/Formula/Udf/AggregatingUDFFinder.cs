/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI.HSSF.Record.Formula.Udf
{
    using System;
    using NPOI.HSSF.Record.Formula.Functions;

    /**
     * Collects Add-in libraries and VB macro functions toGether into one UDF Finder
     *
     * @author PUdalau
     */
    public class AggregatingUDFFinder : UDFFinder
    {

        private UDFFinder[] _usedToolPacks=new UDFFinder[20];

        public AggregatingUDFFinder(params UDFFinder[] usedToolPacks)
        {
            Array.Copy(usedToolPacks, _usedToolPacks, usedToolPacks.Length);
        }

        /**
         * Returns executor by specified name. Returns <code>null</code> if
         * function isn't Contained by any registered tool pack.
         *
         * @param name Name of function.
         * @return Function executor. <code>null</code> if not found
         */
        public override FreeRefFunction FindFunction(String name)
        {
            FreeRefFunction evaluatorForFunction;
            foreach (UDFFinder pack in _usedToolPacks)
            {
                evaluatorForFunction = pack.FindFunction(name);
                if (evaluatorForFunction != null)
                {
                    return evaluatorForFunction;
                }
            }
            return null;
        }
    }
}

