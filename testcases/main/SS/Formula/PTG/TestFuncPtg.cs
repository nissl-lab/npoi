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

namespace TestCases.SS.Formula.PTG
{

    using NUnit.Framework;
    using NPOI.SS.Formula.PTG;
    using TestCases.HSSF.Record;

    /**
     * Make sure the FuncPtg performs as expected
     *
     * @author Danny Mui (dmui at apache dot org)
     */
    [TestFixture]
    public class TestFuncPtg
    {
        [Test]
        public void TestRead()
        {
            // This function index represents the LEN() function
            byte[] fakeData = { 0x20, 0x00, };

            FuncPtg ptg = FuncPtg.Create(TestcaseRecordInputStream.CreateLittleEndian(fakeData));
            Assert.AreEqual(0x20, ptg.FunctionIndex, "Len formula index is not 32(20H)");
            Assert.AreEqual(1, ptg.NumberOfOperands, "Number of operands in the len formula");
            Assert.AreEqual("LEN", ptg.Name, "Function Name");
            Assert.AreEqual(3, ptg.Size, "Ptg Size");
        }
        [Test]
        public void TestNumberOfOperands()
        {
            FuncPtg funcPtg = FuncPtg.Create(27); // ROUND() - takes 2 args
            Assert.AreEqual(2, funcPtg.NumberOfOperands);
            Assert.AreEqual("ROUND", funcPtg.Name);
        }
    }

}