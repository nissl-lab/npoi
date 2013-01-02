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
     * Default UDF Finder - for Adding your own user defined functions.
     *
     * @author PUdalau
     */
    public class DefaultUDFFinder : UDFFinder
    {
        private Dictionary<String, FreeRefFunction> _functionsByName;

        public DefaultUDFFinder(String[] functionNames, FreeRefFunction[] functionImpls)
        {
            int nFuncs = functionNames.Length;
            if (functionImpls.Length != nFuncs)
            {
                throw new ArgumentException(
                        "Mismatch in number of function names and implementations");
            }
            Dictionary<String, FreeRefFunction> m = new Dictionary<String, FreeRefFunction>(nFuncs * 3 / 2);
            for (int i = 0; i < functionImpls.Length; i++)
            {
                m[functionNames[i].ToUpper()]= functionImpls[i];
            }
            _functionsByName = m;
        }

        public override FreeRefFunction FindFunction(String name)
        {
            if (!_functionsByName.ContainsKey(name.ToUpper()))
                return null;

            return _functionsByName[name.ToUpper()];
        }
    }

}