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

using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
using System;
using NPOI.XSSF.Util;

namespace TestCases.XSSF.Util
{

    [TestFixture]
    public class TestCTColComparator
    {
        [Test]
        public void TestCompare()
        {
            CTColComparator comparator = new CTColComparator();
            CT_Col o1 = new CT_Col();
            o1.min = 1;
            o1.max = 10;
            CT_Col o2 = new CT_Col();
            o2.min = 11;
            o2.max = 12;
            Assert.AreEqual(-1, comparator.Compare(o1, o2));
            CT_Col o3 = new CT_Col();
            o3.min = 5;
            o3.max = 8;
            CT_Col o4 = new CT_Col();
            o4.min = 5;
            o4.max = 80;
            Assert.AreEqual(-1, comparator.Compare(o3, o4));
        }
        [Test]
        public void TestArraysSort()
        {
            CTColComparator comparator = new CTColComparator();
            CT_Col o1 = new CT_Col();
            o1.min = 1;
            o1.max = 10;
            CT_Col o2 = new CT_Col();
            o2.min = 11;
            o2.max = 12;
            Assert.AreEqual(-1, comparator.Compare(o1, o2));
            CT_Col o3 = new CT_Col();
            o3.min = 5;
            o3.max = 80;
            CT_Col o4 = new CT_Col();
            o4.min = 5;
            o4.max = 8;
            Assert.AreEqual(1, comparator.Compare(o3, o4));
            CT_Col[] cols = new CT_Col[4];
            cols[0] = o1;
            cols[1] = o2;
            cols[2] = o3;
            cols[3] = o4;
            Assert.AreEqual((uint)80, cols[2].max);
            Assert.AreEqual((uint)8, cols[3].max);
            Array.Sort(cols, comparator);
            Assert.AreEqual((uint)12, cols[3].max);
            Assert.AreEqual((uint)8, cols[1].max);
            Assert.AreEqual((uint)80, cols[2].max);
        }
    }
}

