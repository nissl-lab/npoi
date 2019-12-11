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

namespace TestCases.SS.UserModel
{
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using System;

    /**
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestIndexedColors
    {
        [Test]
        public void FromInt()
        {
            int[] illegalIndices = { -1, 0, 27, 65 };
            foreach (int index in illegalIndices)
            {
                try
                {
                    IndexedColors.FromInt(index);
                    Assert.Fail("Expected ArgumentException: " + index);
                }
                catch (ArgumentException)
                {
                    // expected
                }
            }
            Assert.AreEqual(IndexedColors.Black, IndexedColors.FromInt(8));
            Assert.AreEqual(IndexedColors.Gold, IndexedColors.FromInt(51));
            Assert.AreEqual(IndexedColors.Automatic, IndexedColors.FromInt(64));
        }


        [Test]
        public void GetIndex()
        {
            Assert.AreEqual(51, IndexedColors.Gold.Index);
        }

        [Test]
        public void Index()
        {
            Assert.AreEqual(51, IndexedColors.Gold.Index);
        }
    }
}