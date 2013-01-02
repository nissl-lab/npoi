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

namespace TestCases.HSSF.Util
{
    using System;

    using NUnit.Framework;
    using NPOI.Util;
    using NPOI.HSSF.Util;

    /**
     * Tests the {@link RKUtil} class.
     */
    [TestFixture]
    public class TestRKUtil
    {

        /**
         * Check we can decode correctly.
         */
        [Test]
        public void TestDecode()
        {

            int[] values = { 1074266112, 1081384961, 1081397249, 
				0x3FF00000, 0x405EC001, 0x02F1853A, 0x02F1853B, unchecked((int)0xFCDD699A),
		};
            double[] rvalues = { 3.0, 3.3, 3.33,
				1, 1.23, 12345678, 123456.78, -13149594, 
		};

            for (int j = 0; j < values.Length; j++)
            {

                int intBits = values[j];
                double expectedValue = rvalues[j];
                double actualValue = RKUtil.DecodeNumber(intBits);
                if (expectedValue != actualValue)
                {
                    throw new AssertionException("0x" + StringUtil.ToHexString(intBits)
                            + " should decode to " + expectedValue + " but got " + actualValue);
                }
            }
        }
    }

}