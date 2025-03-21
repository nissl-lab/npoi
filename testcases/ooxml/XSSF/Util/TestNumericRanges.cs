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

using NPOI.XSSF.Util;
using NUnit.Framework;using NUnit.Framework.Legacy;
namespace TestCases.XSSF.Util
{

    [TestFixture]
    public class TestNumericRanges
    {
        [Test]
        public void TestGetOverlappingType()
        {
            long[] r1 = { 3, 8 };
            long[] r2 = { 6, 11 };
            long[] r3 = { 1, 5 };
            long[] r4 = { 2, 20 };
            long[] r5 = { 5, 6 };
            long[] r6 = { 20, 23 };
            ClassicAssert.AreEqual(NumericRanges.OVERLAPS_1_MINOR, NumericRanges.GetOverlappingType(r1, r2));
            ClassicAssert.AreEqual(NumericRanges.OVERLAPS_2_MINOR, NumericRanges.GetOverlappingType(r1, r3));
            ClassicAssert.AreEqual(NumericRanges.OVERLAPS_2_WRAPS, NumericRanges.GetOverlappingType(r1, r4));
            ClassicAssert.AreEqual(NumericRanges.OVERLAPS_1_WRAPS, NumericRanges.GetOverlappingType(r1, r5));
            ClassicAssert.AreEqual(NumericRanges.NO_OVERLAPS, NumericRanges.GetOverlappingType(r1, r6));
        }
        [Test]
        public void TestGetOverlappingRange()
        {
            long[] r1 = { 3, 8 };
            long[] r2 = { 6, 11 };
            long[] r3 = { 1, 5 };
            long[] r4 = { 2, 20 };
            long[] r5 = { 5, 6 };
            long[] r6 = { 20, 23 };
            ClassicAssert.AreEqual(6, NumericRanges.GetOverlappingRange(r1, r2)[0]);
            ClassicAssert.AreEqual(8, NumericRanges.GetOverlappingRange(r1, r2)[1]);
            ClassicAssert.AreEqual(3, NumericRanges.GetOverlappingRange(r1, r3)[0]);
            ClassicAssert.AreEqual(5, NumericRanges.GetOverlappingRange(r1, r3)[1]);
            ClassicAssert.AreEqual(3, NumericRanges.GetOverlappingRange(r1, r4)[0]);
            ClassicAssert.AreEqual(8, NumericRanges.GetOverlappingRange(r1, r4)[1]);
            ClassicAssert.AreEqual(5, NumericRanges.GetOverlappingRange(r1, r5)[0]);
            ClassicAssert.AreEqual(6, NumericRanges.GetOverlappingRange(r1, r5)[1]);
            ClassicAssert.AreEqual(-1, NumericRanges.GetOverlappingRange(r1, r6)[0]);
            ClassicAssert.AreEqual(-1, NumericRanges.GetOverlappingRange(r1, r6)[1]);
        }

    }

}