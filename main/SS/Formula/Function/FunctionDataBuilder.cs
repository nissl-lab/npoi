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


    /**
     * Temporarily collects <c>FunctionMetadata</c> instances for creation of a
     * <c>FunctionMetadataRegistry</c>.
     * 
     * @author Josh Micich
     */
    class FunctionDataBuilder
    {
        private int _maxFunctionIndex;
        private Hashtable _functionDataByName;
        private Hashtable _functionDataByIndex;
        /** stores indexes of all functions with footnotes (i.e. whose definitions might Change) */
        private ArrayList _mutatingFunctionIndexes;

        public FunctionDataBuilder(int sizeEstimate)
        {
            _maxFunctionIndex = -1;
            _functionDataByName = new Hashtable(sizeEstimate * 3 / 2);
            _functionDataByIndex = new Hashtable(sizeEstimate * 3 / 2);
            _mutatingFunctionIndexes = new ArrayList();
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
            prevFM = (FunctionMetadata)_functionDataByName[functionName];
            if (prevFM != null)
            {
                if (!hasFootnote || !_mutatingFunctionIndexes.Contains(indexKey))
                {
                    throw new Exception("Multiple entries for function name '" + functionName + "'");
                }
                _functionDataByIndex.Remove(prevFM.Index);
            }
            prevFM = (FunctionMetadata)_functionDataByIndex[indexKey];
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