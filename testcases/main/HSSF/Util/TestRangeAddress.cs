
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
    using NUnit.Framework;

    /**
     * Tests the Range Address Utility Functionality
     * @author Danny Mui (danny at muibros.com)
     */
    [TestFixture]
    public class TestRangeAddress
    {
        public TestRangeAddress()
        {

        }

        [Test]
        public void TestReferenceParse()
        {
            String reference = "Sheet2!$A$1:$C$3";
            RangeAddress ra = new RangeAddress(reference);

            Assert.AreEqual("Sheet2!A1:C3", ra.Address);

        }
    }
}
