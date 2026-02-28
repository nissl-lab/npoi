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


using NPOI.SS.Util;
using NPOI.Util;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.SS.Util
{

    /// <summary>
    /// <para>
    /// Tests CellRangeUtil.
    /// </para>
    /// <para>
    /// <see cref="NPOI.SS.Util.CellRangeUtil" />
    /// </para>
    /// </summary>
    [TestFixture]
    public sealed class TestCellRangeUtil
    {

        private static CellRangeAddress A1 = new CellRangeAddress(0, 0, 0, 0);
        private static CellRangeAddress B1 = new CellRangeAddress(0, 0, 1, 1);
        private static CellRangeAddress A2 = new CellRangeAddress(1, 1, 0, 0);
        private static CellRangeAddress B2 = new CellRangeAddress(1, 1, 1, 1);
        private static CellRangeAddress A1_B2 = new CellRangeAddress(0, 1, 0, 1);
        private static CellRangeAddress A1_B1 = new CellRangeAddress(0, 0, 0, 1);
        private static CellRangeAddress A1_A2 = new CellRangeAddress(0, 1, 0, 0);

        [Test]
        public void TestMergeCellRanges()
        {
            // Note that the order of the output array elements does not matter
            // And that there may be more than one valid outputs for a given input. Any valid output is accepted.
            // POI should use a strategy that is consistent and predictable (it currently is not).

            // Fully mergeable
            //    A B
            //  1 x x   A1,A2,B1,B2 --> A1:B2
            //  2 x x
            AssertCellRangesEqual(AsArray(A1_B2), Merge(A1, B1, A2, B2));
            AssertCellRangesEqual(AsArray(A1_B2), Merge(A1, B2, A2, B1));

            // Partially mergeable: multiple possible merges
            //    A B
            //  1 x x   A1,A2,B1 --> A1:B1,A2 or A1:A2,B1
            //  2 x 
            AssertCellRangesEqual(AsArray(A1_B1, A2), Merge(A1, B1, A2));
            AssertCellRangesEqual(AsArray(A1_A2, B1), Merge(A2, A1, B1));
            AssertCellRangesEqual(AsArray(A1_B1, A2), Merge(B1, A2, A1));

            // Not mergeable
            //    A B
            //  1 x     A1,B2 --> A1,B2
            //  2   x
            AssertCellRangesEqual(AsArray(A1, B2), Merge(A1, B2));
            AssertCellRangesEqual(AsArray(B2, A1), Merge(B2, A1));
        }

        private void AssertCellRangesEqual(CellRangeAddress[] a, CellRangeAddress[] b)
        {
            ClassicAssert.AreEqual(GetCellAddresses(a), GetCellAddresses(b));
            CollectionAssert.AreEquivalent(a, b);
            //assertArrayEquals(a, b);
        }

        private static HashSet<CellAddress> GetCellAddresses(CellRangeAddress[] ranges)
        {
            HashSet<CellAddress> set = new HashSet<CellAddress>();
            foreach(CellRangeAddress range in ranges)
            {
                foreach(var ca in range)
                {
                    set.Add(ca);
                }
            }
            return set;
        }

        private static CellRangeAddress[] AsArray(params CellRangeAddress[] ts)
        {
            return ts;
        }

        private static CellRangeAddress[] Merge(params CellRangeAddress[] ranges)
        {
            return CellRangeUtil.MergeCellRanges(ranges);
        }
    }
}


