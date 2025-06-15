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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.Formula.PTG
{
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    [TestFixture]
    public class TestAbstractFunctionPtg
    {
        [Test]
        public void TestConstructor()
        {
            FunctionPtg ptg = new FunctionPtg(1, 2, null, 255);
            ClassicAssert.AreEqual(1, ptg.FunctionIndex);
            ClassicAssert.AreEqual(2, ptg.DefaultOperandClass);
            ClassicAssert.AreEqual(255, ptg.NumberOfOperands);
        }

        [Test]
        public void TestInvalidFunctionIndex()
        {
            ClassicAssert.Throws<RuntimeException>(()=>{
                new FunctionPtg(40000, 2, null, 255);
            });
            
        }

        [Test]
        public void TestInvalidRuntimeClass()
        {
            ClassicAssert.Throws<RuntimeException>(()=>{
                new FunctionPtg(1, 300, null, 255);
            });
        }

        private class FunctionPtg : AbstractFunctionPtg
        {

            internal FunctionPtg(int functionIndex, int pReturnClass,
                    byte[] paramTypes, int nParams)
                    : base(functionIndex, pReturnClass, paramTypes, nParams)
            {

            }

            public override int Size
            {
                get
                {
                    return 0;
                }
            }

            public override void Write(ILittleEndianOutput out1)
            {

            }
        }
    }
}


