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

namespace NPOI.SS.Formula.Function
{
    using System;
    using System.Collections;
    using System.Collections.Generic;

    /**
     * Temporarily collects <c>FunctionMetadata</c> instances for creation of a
     * <c>FunctionMetadataRegistry</c>.
     * 
     * @author Josh Micich
     */
    class FunctionDataBuilder
    {
        private int _maxFunctionIndex;

        private readonly Dictionary<string, FunctionMetadata> _functionDataByName;
        private readonly Dictionary<int, FunctionMetadata> _functionDataByIndex;
        /** stores indexes of all functions with footnotes (i.e. whose definitions might Change) */
        private readonly HashSet<int> _mutatingFunctionIndexes;

        public FunctionDataBuilder(int sizeEstimate)
        {
            _maxFunctionIndex = -1;
            _functionDataByName = new Dictionary<string, FunctionMetadata>(sizeEstimate * 3 / 2);
            _functionDataByIndex = new Dictionary<int, FunctionMetadata>(sizeEstimate * 3 / 2);
            _mutatingFunctionIndexes = new HashSet<int>();
        }

        public void Add(int functionIndex, String functionName, int minParams, int maxParams,
                byte returnClassCode, byte[] parameterClassCodes, bool hasFootnote)
        {
            FunctionMetadata fm = new FunctionMetadata(functionIndex, functionName, minParams, maxParams,
                    returnClassCode, parameterClassCodes);

            int indexKey = functionIndex;


            if (functionIndex > _maxFunctionIndex)
            {
                _maxFunctionIndex = functionIndex;
            }
            // allow function definitions to Change only if both previous and the new items have footnotes
            FunctionMetadata prevFM;
            _functionDataByName.TryGetValue(functionName, out prevFM);
            if (prevFM != null)
            {
                if (!hasFootnote || !_mutatingFunctionIndexes.Contains(indexKey))
                {
                    throw new Exception("Multiple entries for function name '" + functionName + "'");
                }
                _functionDataByIndex.Remove(prevFM.Index);
            }
            //prevFM = _functionDataByIndex[indexKey];
            _functionDataByIndex.TryGetValue(indexKey, out prevFM);
            if (prevFM != null)
            {
                if (!hasFootnote || !_mutatingFunctionIndexes.Contains(indexKey))
                {
                    throw new Exception("Multiple entries for function index (" + functionIndex + ")");
                }
                _functionDataByName.Remove(prevFM.Name);
            }
            if (hasFootnote)
            {
                _mutatingFunctionIndexes.Add(indexKey);
            }
            _functionDataByIndex[indexKey]=fm;
            _functionDataByName[functionName]=fm;
        }

        public FunctionMetadataRegistry Build()
        {

            FunctionMetadata[] jumbledArray = new FunctionMetadata[_functionDataByName.Count];
            IEnumerator values = _functionDataByName.Values.GetEnumerator();
            FunctionMetadata[] fdIndexArray = new FunctionMetadata[_maxFunctionIndex + 1];
            while (values.MoveNext())
            {
                FunctionMetadata fd = (FunctionMetadata)values.Current;
                fdIndexArray[fd.Index] = fd;
            }

            return new FunctionMetadataRegistry(fdIndexArray, _functionDataByName);
        }
    }
}