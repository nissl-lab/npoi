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
     * Allows clients to Get <c>FunctionMetadata</c> instances for any built-in function of Excel.
     * 
     * @author Josh Micich
     */
    public class FunctionMetadataRegistry
    {
        /**
         * The name of the IF function (i.e. "IF").  Extracted as a constant for clarity.
         */
        public const String FUNCTION_NAME_IF = "IF";

        public const int FUNCTION_INDEX_IF = 1;
        public const short FUNCTION_INDEX_SUM = 4;
        public const int FUNCTION_INDEX_CHOOSE = 100;
	    public const short FUNCTION_INDEX_INDIRECT = 148;
        public const short FUNCTION_INDEX_EXTERNAL = 255;
        private static FunctionMetadataRegistry _instance;

        private FunctionMetadata[] _functionDataByIndex;
        private Hashtable _functionDataByName;

        private static FunctionMetadataRegistry GetInstance()
        {
            if (_instance == null)
            {
                _instance = FunctionMetadataReader.CreateRegistry();
            }
            return _instance;
        }

        /* package */
        public FunctionMetadataRegistry(FunctionMetadata[] functionDataByIndex, Hashtable functionDataByName)
        {
            _functionDataByIndex = functionDataByIndex;
            _functionDataByName = functionDataByName;
        }

        /* package */
        public ICollection GetAllFunctionNames()
        {
            return _functionDataByName.Keys;
        }


        public static FunctionMetadata GetFunctionByIndex(int index)
        {
            return GetInstance().GetFunctionByIndexInternal(index);
        }

        private FunctionMetadata GetFunctionByIndexInternal(int index)
        {
            return _functionDataByIndex[index];
        }
        /**
         * Resolves a built-in function index. 
         * @param name uppercase function name
         * @return a negative value if the function name is not found.
         * This typically occurs for external functions.
         */
        public static short LookupIndexByName(String name)
        {
            FunctionMetadata fd = GetInstance().GetFunctionByNameInternal(name);
            if (fd == null)
            {
                return -1;
            }
            return (short)fd.Index;
        }

        private FunctionMetadata GetFunctionByNameInternal(String name)
        {
            return (FunctionMetadata)_functionDataByName[name];
        }


        public static FunctionMetadata GetFunctionByName(String name)
        {
            return GetInstance().GetFunctionByNameInternal(name);
        }
    }
}