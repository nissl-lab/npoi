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
    using System.Text;
    using System.Collections;
    using NPOI.HSSF.Util;

    using NUnit.Framework;using NUnit.Framework.Legacy;
    using System.Collections.Generic;

    /**
* @author Nick Burch
*/
    [TestFixture]
    public class TestHSSFColor
    {
        [Test]
        public void TestBasics()
        {
            ClassicAssert.IsNotNull(typeof(HSSFColor.Yellow));
            ClassicAssert.IsTrue(HSSFColor.Yellow.Index > 0);
            ClassicAssert.IsTrue(HSSFColor.Yellow.Index2 > 0);
        }
        [Test]
        public void TestContents()
        {
            ClassicAssert.AreEqual(3, HSSFColor.Yellow.Triplet.Length);
            ClassicAssert.AreEqual(255, HSSFColor.Yellow.Triplet[0]);
            ClassicAssert.AreEqual(255, HSSFColor.Yellow.Triplet[1]);
            ClassicAssert.AreEqual(0, HSSFColor.Yellow.Triplet[2]);

            ClassicAssert.AreEqual("FFFF:FFFF:0", HSSFColor.Yellow.HexString);
        }
        [Test]
        public void TestTrippletHash()
        {
            Dictionary<String, HSSFColor> tripplets = HSSFColor.GetTripletHash();

            ClassicAssert.AreEqual(
                    typeof(HSSFColor.Maroon),
                    tripplets[HSSFColor.Maroon.HexString].GetType()
            );
            ClassicAssert.AreEqual(
                    typeof(HSSFColor.Yellow),
                    tripplets[HSSFColor.Yellow.HexString].GetType()
            );
        }
    }
}
