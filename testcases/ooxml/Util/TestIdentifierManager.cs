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
using NPOI.Util;
using NUnit.Framework;
using System;
namespace TestCases.Util
{
    [TestFixture]
    public class TestIdentifierManager
    {
        [Test]
        public void TestBasic()
        {
            IdentifierManager manager = new IdentifierManager(0L, 100L);
            Assert.AreEqual(101L, manager.GetRemainingIdentifiers());
            Assert.AreEqual(0L, manager.ReserveNew());
            Assert.AreEqual(100L, manager.GetRemainingIdentifiers());
            Assert.AreEqual(1L, manager.Reserve(0L));
            Assert.AreEqual(99L, manager.GetRemainingIdentifiers());
        }
        [Test]
        public void TestLongLimits()
        {
            long min = IdentifierManager.MIN_ID;
            long max = IdentifierManager.MAX_ID;
            IdentifierManager manager = new IdentifierManager(min, max);
            Assert.IsTrue(max - min + 1 > 0, "Limits lead to a long variable overflow");
            Assert.IsTrue(manager.GetRemainingIdentifiers() > 0, "Limits lead to a long variable overflow");
            Assert.AreEqual(min, manager.ReserveNew());
            Assert.AreEqual(max, manager.Reserve(max));
            Assert.AreEqual(max - min - 1, manager.GetRemainingIdentifiers());
            manager.Release(max);
            manager.Release(min);
        }
        [Test]
        public void TestReserve()
        {
            IdentifierManager manager = new IdentifierManager(10L, 30L);
            Assert.AreEqual(12L, manager.Reserve(12L));
            long reserve = manager.Reserve(12L);
            Assert.IsFalse(reserve == 12L, "Same id must be reserved twice!");
            Assert.IsTrue(manager.Release(12L));
            Assert.IsTrue(manager.Release(reserve));
            Assert.IsFalse(manager.Release(12L));
            Assert.IsFalse(manager.Release(reserve));

            manager = new IdentifierManager(0L, 2L);
            Assert.AreEqual(0L, manager.Reserve(0L));
            Assert.AreEqual(1L, manager.Reserve(1L));
            Assert.AreEqual(2L, manager.Reserve(2L));
            try
            {
                manager.Reserve(0L);
                Assert.Fail("Exception expected");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
            try
            {
                manager.Reserve(1L);
                Assert.Fail("Exception expected");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
            try
            {
                manager.Reserve(2L);
                Assert.Fail("Exception expected");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
        }
        [Test]
        public void TestReserveNew()
        {
            IdentifierManager manager = new IdentifierManager(10L, 12L);
            Assert.AreEqual(10L, manager.ReserveNew());
            Assert.AreEqual(11L, manager.ReserveNew());
            Assert.AreEqual(12L, manager.ReserveNew());
            try
            {
                manager.ReserveNew();
                Assert.Fail("InvalidOperationException expected");
            }
            catch (InvalidOperationException)
            {
                // expected
            }
        }
        [Test]
        public void TestRelease()
        {
            IdentifierManager manager = new IdentifierManager(10L, 20L);
            Assert.AreEqual(10L, manager.Reserve(10L));
            Assert.AreEqual(11L, manager.Reserve(11L));
            Assert.AreEqual(12L, manager.Reserve(12L));
            Assert.AreEqual(13L, manager.Reserve(13L));
            Assert.AreEqual(14L, manager.Reserve(14L));

            Assert.IsTrue(manager.Release(10L));
            Assert.AreEqual(10L, manager.Reserve(10L));
            Assert.IsTrue(manager.Release(10L));

            Assert.IsTrue(manager.Release(11L));
            Assert.AreEqual(11L, manager.Reserve(11L));
            Assert.IsTrue(manager.Release(11L));
            Assert.IsFalse(manager.Release(11L));
            Assert.IsFalse(manager.Release(10L));

            Assert.AreEqual(10L, manager.Reserve(10L));
            Assert.AreEqual(11L, manager.Reserve(11L));
            Assert.IsTrue(manager.Release(12L));
        }
    }


}