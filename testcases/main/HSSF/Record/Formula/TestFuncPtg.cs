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

namespace TestCases.HSSF.Record.Formula
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Formula;


    /**
     * Make sure the FuncPtg performs as expected
     *
     * @author Danny Mui (dmui at apache dot org)
     */
    [TestClass]
    public class TestFuncPtg
    {
        [TestMethod]
        public void TestRead()
        {
            // This ptg represents a LEN function extracted from excel
            byte[] fakeData = {
            0x20,  //function index
            0,
        };

            FuncPtg ptg = new FuncPtg(TestcaseRecordInputStream.CreateWithFakeSid(fakeData));
            Assert.AreEqual(0x20, ptg.GetFunctionIndex(), "Len formula index is1 not 32(20H)");
            Assert.AreEqual(1, ptg.NumberOfOperands, "Number of operands in the len formula");
            Assert.AreEqual("LEN", ptg.Name, "Function Name");
            Assert.AreEqual(3, ptg.Size, "Ptg Size");
        }
        [TestMethod]
        public void TestClone()
        {
            FuncPtg funcPtg = new FuncPtg(27); // ROUND() - takes 2 args

            FuncPtg clone = (FuncPtg)funcPtg.Clone();
            if (clone.NumberOfOperands == 0)
            {
                Assert.Fail("clone() did copy field numberOfOperands");
            }
            Assert.AreEqual(2, clone.NumberOfOperands);
            Assert.AreEqual("ROUND", clone.Name);
        }
    }
}