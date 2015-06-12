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

namespace NPOI.SS.Formula.Udf
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS.Formula.Functions;
    /**
     * A UDFFinder that can retrieve functions both by name and by fake index.
     *
     * @author Yegor Kozlov
     */
    public class IndexedUDFFinder : AggregatingUDFFinder
    {
        private Dictionary<int, String> _funcMap;

        public IndexedUDFFinder(params UDFFinder[] usedToolPacks)
            : base(usedToolPacks)
        {

            _funcMap = new Dictionary<int, String>();
        }

        public override FreeRefFunction FindFunction(String name)
        {
            FreeRefFunction func = base.FindFunction(name);
            if (func != null)
            {
                int idx = GetFunctionIndex(name);
                _funcMap[idx] = name;
            }
            return func;
        }

        public String GetFunctionName(int idx)
        {
            return _funcMap[idx];
        }

        public int GetFunctionIndex(String name)
        {
            return name.GetHashCode();
        }
    }
}





