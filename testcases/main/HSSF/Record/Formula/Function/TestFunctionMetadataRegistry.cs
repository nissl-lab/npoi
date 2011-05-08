/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record.Formula.Function
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record.Formula.Function;
    /**
     * 
     * @author Josh Micich
     */
    public class TestFunctionMetadataRegistry
    {

        public void TestWellKnownFunctions()
        {
            ConfirmFunction(0, "COUNT");
            ConfirmFunction(1, "IF");

        }

        private static void ConfirmFunction(int index, String funcName)
        {
            FunctionMetadata fm;
            fm = FunctionMetadataRegistry.GetFunctionByIndex(index);
            Assert.IsNotNull(fm);
            Assert.AreEqual(funcName, fm.Name);

            fm = FunctionMetadataRegistry.GetFunctionByName(funcName);
            Assert.IsNotNull(fm);
            Assert.AreEqual(index, fm.Index);
        }
    }
}