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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.Text;
using System.Collections.Generic;

using NUnit.Framework;
using NPOI.Util;

namespace TestCases.Util
{
    /// <summary>
    /// Summary description for TestBitField
    /// </summary>
    [TestFixture]
    public class TestBitField
    {
        private static BitField bf_multi = BitFieldFactory.GetInstance(0x3F80);
        private static BitField bf_single = BitFieldFactory.GetInstance(0x4000);

        public TestBitField()
        {

        }

        /// <summary>
        /// Tests the get value.
        /// </summary>
        [Test]
        public void TestGetValue()
        {
            Assert.AreEqual(bf_multi.GetValue(-1), 127);
            Assert.AreEqual(bf_multi.GetValue(0), 0);
            Assert.AreEqual(bf_single.GetValue(-1), 1);
            Assert.AreEqual(bf_single.GetValue(0), 0);
        }
        /// <summary>
        /// Tests the get short value.
        /// </summary>
        [Test]
        public void TestGetShortValue()
        {
            Assert.AreEqual(bf_multi.GetShortValue((short)-1), (short)127);
            Assert.AreEqual(bf_multi.GetShortValue((short)0), (short)0);
            Assert.AreEqual(bf_single.GetShortValue((short)-1), (short)1);
            Assert.AreEqual(bf_single.GetShortValue((short)0), (short)0);
        }
        /// <summary>
        /// Tests the get raw value.
        /// </summary>
        [Test]
        public void TestGetRawValue()
        {
            Assert.AreEqual(bf_multi.GetRawValue(-1), 0x3F80);
            Assert.AreEqual(bf_multi.GetRawValue(0), 0);
            Assert.AreEqual(bf_single.GetRawValue(-1), 0x4000);
            Assert.AreEqual(bf_single.GetRawValue(0), 0);
        }
        /// <summary>
        /// Tests the get short raw value.
        /// </summary>
        [Test]
        public void TestGetShortRawValue()
        {
            Assert.AreEqual(bf_multi.GetShortRawValue((short)-1),
                         (short)0x3F80);
            Assert.AreEqual(bf_multi.GetShortRawValue((short)0), (short)0);
            Assert.AreEqual(bf_single.GetShortRawValue((short)-1),
                         (short)0x4000);
            Assert.AreEqual(bf_single.GetShortRawValue((short)0), (short)0);
        }
        /// <summary>
        /// Tests the is set.
        /// </summary>
        [Test]
        public void TestIsSet()
        {
            Assert.IsTrue(!bf_multi.IsSet(0));
            for (int j = 0x80; j <= 0x3F80; j += 0x80)
            {
                Assert.IsTrue(bf_multi.IsSet(j));
            }
            Assert.IsTrue(!bf_single.IsSet(0));
            Assert.IsTrue(bf_single.IsSet(0x4000));
        }
        /// <summary>
        /// Tests the is all set.
        /// </summary>
        [Test]
        public void TestIsAllSet()
        {
            for (int j = 0; j < 0x3F80; j += 0x80)
            {
                Assert.IsTrue(!bf_multi.IsAllSet(j));
            }
            Assert.IsTrue(bf_multi.IsAllSet(0x3F80));
            Assert.IsTrue(!bf_single.IsAllSet(0));
            Assert.IsTrue(bf_single.IsAllSet(0x4000));
        }
        /// <summary>
        /// Tests the set value.
        /// </summary>
        [Test]
        public void TestSetValue()
        {
            for (int j = 0; j < 128; j++)
            {
                Assert.AreEqual(bf_multi.GetValue(bf_multi.SetValue(0, j)), j);
                Assert.AreEqual(bf_multi.SetValue(0, j), j << 7);
            }

            // verify that excess bits are stripped off
            Assert.AreEqual(bf_multi.SetValue(0x3f80, 128), 0);
            for (int j = 0; j < 2; j++)
            {
                Assert.AreEqual(bf_single.GetValue(bf_single.SetValue(0, j)), j);
                Assert.AreEqual(bf_single.SetValue(0, j), j << 14);
            }

            // verify that excess bits are stripped off
            Assert.AreEqual(bf_single.SetValue(0x4000, 2), 0);
        }
        /// <summary>
        /// Tests the set short value.
        /// </summary>
        [Test]
        public void TestSetShortValue()
        {
            for (int j = 0; j < 128; j++)
            {
                Assert.AreEqual(bf_multi
                    .GetShortValue(bf_multi
                        .SetShortValue((short)0, (short)j)), (short)j);
                Assert.AreEqual(bf_multi.SetShortValue((short)0, (short)j),
                             (short)(j << 7));
            }

            // verify that excess bits are stripped off
            Assert.AreEqual(bf_multi.SetShortValue((short)0x3f80, (short)128),
                         (short)0);
            for (int j = 0; j < 2; j++)
            {
                Assert.AreEqual(bf_single
                    .GetShortValue(bf_single
                        .SetShortValue((short)0, (short)j)), (short)j);
                Assert.AreEqual(bf_single.SetShortValue((short)0, (short)j),
                             (short)(j << 14));
            }

            // verify that excess bits are stripped off
            Assert.AreEqual(bf_single.SetShortValue((short)0x4000, (short)2),
                         (short)0);
        }
        /// <summary>
        /// Tests the byte.
        /// </summary>
        [Test]
        public void TestByte()
        {
            Assert.AreEqual(1, BitFieldFactory.GetInstance(1).SetByteBoolean((byte)0, true));
            Assert.AreEqual(2, BitFieldFactory.GetInstance(2).SetByteBoolean((byte)0, true));
            Assert.AreEqual(4, BitFieldFactory.GetInstance(4).SetByteBoolean((byte)0, true));
            Assert.AreEqual(8, BitFieldFactory.GetInstance(8).SetByteBoolean((byte)0, true));
            Assert.AreEqual(16, BitFieldFactory.GetInstance(16).SetByteBoolean((byte)0, true));
            Assert.AreEqual(32, BitFieldFactory.GetInstance(32).SetByteBoolean((byte)0, true));
            Assert.AreEqual(64, BitFieldFactory.GetInstance(64).SetByteBoolean((byte)0, true));
            Assert.AreEqual(128, BitFieldFactory.GetInstance(128).SetByteBoolean((byte)0, true));
            Assert.AreEqual(0, BitFieldFactory.GetInstance(1).SetByteBoolean((byte)1, false));
            Assert.AreEqual(0, BitFieldFactory.GetInstance(2).SetByteBoolean((byte)2, false));
            Assert.AreEqual(0, BitFieldFactory.GetInstance(4).SetByteBoolean((byte)4, false));
            Assert.AreEqual(0, BitFieldFactory.GetInstance(8).SetByteBoolean((byte)8, false));
            Assert.AreEqual(0, BitFieldFactory.GetInstance(16).SetByteBoolean((byte)16, false));
            Assert.AreEqual(0, BitFieldFactory.GetInstance(32).SetByteBoolean((byte)32, false));
            Assert.AreEqual(0, BitFieldFactory.GetInstance(64).SetByteBoolean((byte)64, false));
            Assert.AreEqual(0, BitFieldFactory.GetInstance(127).SetByteBoolean((byte)127, false));
            Assert.AreEqual(254, BitFieldFactory.GetInstance(1).SetByteBoolean((byte)254, false));
            byte clearedBit = BitFieldFactory.GetInstance(0x40).SetByteBoolean(unchecked((byte)-63), false);

            Assert.AreEqual(false, BitFieldFactory.GetInstance(0x40).IsSet(clearedBit));
        }
        /// <summary>
        /// Tests the clear.
        /// </summary>
        [Test]
        public void TestClear()
        {
            Assert.AreEqual(bf_multi.Clear(-1), unchecked((Int32)0xFFFFC07F));
            Assert.AreEqual(bf_single.Clear(-1), unchecked((Int32)0xFFFFBFFF));
        }

        /// <summary>
        /// Tests the clear short.
        /// </summary>
        [Test]
        public void TestClearShort()
        {
            Assert.AreEqual(bf_multi.ClearShort((short)-1), unchecked((short)0xC07F));
            Assert.AreEqual(bf_single.ClearShort((short)-1), unchecked((short)0xBFFF));
        }
        /// <summary>
        /// Tests the set.
        /// </summary>
        [Test]
        public void TestSet()
        {
            Assert.AreEqual(bf_multi.Set(0), 0x3F80);
            Assert.AreEqual(bf_single.Set(0), 0x4000);
        }
        /// <summary>
        /// Tests the set short.
        /// </summary>
        [Test]
        public void TestSetShort()
        {
            Assert.AreEqual(bf_multi.SetShort((short)0), (short)0x3F80);
            Assert.AreEqual(bf_single.SetShort((short)0), (short)0x4000);
        }
        /// <summary>
        /// Tests the set boolean.
        /// </summary>
        [Test]
        public void TestSetBoolean()
        {
            Assert.AreEqual(bf_multi.Set(0), bf_multi.SetBoolean(0, true));
            Assert.AreEqual(bf_single.Set(0), bf_single.SetBoolean(0, true));
            Assert.AreEqual(bf_multi.Clear(-1), bf_multi.SetBoolean(-1, false));
            Assert.AreEqual(bf_single.Clear(-1), bf_single.SetBoolean(-1, false));
        }
        /// <summary>
        /// Tests the set short boolean.
        /// </summary>
        [Test]
        public void TestSetShortBoolean()
        {
            Assert.AreEqual(bf_multi.SetShort((short)0),
                         bf_multi.SetShortBoolean((short)0, true));
            Assert.AreEqual(bf_single.SetShort((short)0),
                         bf_single.SetShortBoolean((short)0, true));
            Assert.AreEqual(bf_multi.ClearShort((short)-1),
                         bf_multi.SetShortBoolean((short)-1, false));
            Assert.AreEqual(bf_single.ClearShort((short)-1),
                         bf_single.SetShortBoolean((short)-1, false));
        }
        [Test]
        public void TestSetLargeValues()
        {
            BitField bf1 = new BitField(0xF), bf2 = new BitField(0xF0000000);
            int a = 0;
            a = bf1.SetValue(a, 9);
            a = bf2.SetValue(a, 9);
            Assert.AreEqual(9, bf1.GetValue(a));
            Assert.AreEqual(9, bf2.GetValue(a));
        }
    }
}
