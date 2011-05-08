
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


namespace TestCases.HSSF.Util
{

    using System;
    using NPOI.HSSF.Util;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /**
     * Tests the RKUtil class.
     */
    [TestClass]
    public class TestRKUtil
    {
        public TestRKUtil()
        {
        }

        /**
         * Check we can Decode correctly.
         */
        [TestMethod]
        public void TestDecode()
        {
            Assert.AreEqual(3.0, RKUtil.DecodeNumber(1074266112), 0.0000001);
            Assert.AreEqual(3.3, RKUtil.DecodeNumber(1081384961), 0.0000001);
            Assert.AreEqual(3.33, RKUtil.DecodeNumber(1081397249), 0.0000001);
        }
    }
}