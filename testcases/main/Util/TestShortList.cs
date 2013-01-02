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

namespace TestCases.Util
{
    using System;

    using NUnit.Framework;
    using NPOI.Util;

    /**
     * Class to Test ShortList
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestShortList
    {
        [Test]
        public void TestConstructors()
        {
            ShortList list = new ShortList();

            Assert.IsTrue(list.IsEmpty());
            list.Add((short)0);
            list.Add((short)1);
            ShortList list2 = new ShortList(list);
            //Assert.AreEqual(list, list2);
            Assert.IsTrue(list.Equals(list2));
            ShortList list3 = new ShortList(2);

            Assert.IsTrue(list3.IsEmpty());
        }
        [Test]
        public void TestAdd()
        {
            ShortList list = new ShortList();
            short[] testArray =
        {
            0, 1, 2, 3, 5
        };

            for (int j = 0; j < testArray.Length; j++)
            {
                list.Add(testArray[j]);
            }
            for (int j = 0; j < testArray.Length; j++)
            {
                Assert.AreEqual(testArray[j], list.Get(j));
            }
            Assert.AreEqual(testArray.Length, list.Count);

            // add at the beginning
            list.Add(0, (short)-1);
            Assert.AreEqual((short)-1, list.Get(0));
            Assert.AreEqual(testArray.Length + 1, list.Count);
            for (int j = 0; j < testArray.Length; j++)
            {
                Assert.AreEqual(testArray[j], list.Get(j + 1));
            }

            // add in the middle
            list.Add(5, (short)4);
            Assert.AreEqual((short)4, list.Get(5));
            Assert.AreEqual(testArray.Length + 2, list.Count);
            for (int j = 0; j < list.Count; j++)
            {
                Assert.AreEqual((short)(j - 1), list.Get(j));
            }

            // add at the end
            list.Add(list.Count, (short)6);
            Assert.AreEqual(testArray.Length + 3, list.Count);
            for (int j = 0; j < list.Count; j++)
            {
                Assert.AreEqual((short)(j - 1), list.Get(j));
            }

            // add past end
            try
            {
                list.Add(list.Count + 1, (short)8);
                Assert.Fail("should have thrown exception");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }

            // Test growth
            list = new ShortList(0);
            for (short j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            Assert.AreEqual(1000, list.Count);
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(j, list.Get(j));
            }
            list = new ShortList(0);
            for (short j = 0; j < 1000; j++)
            {
                list.Add(0, j);
            }
            Assert.AreEqual(1000, list.Count);
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(j, list.Get(999 - j));
            }
        }
        [Test]
        public void TestAddAll()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 5; j++)
            {
                list.Add(j);
            }
            ShortList list2 = new ShortList(0);

            list2.AddAll(list);
            list2.AddAll(list);
            Assert.AreEqual(2 * list.Count, list2.Count);
            for (short j = 0; j < 5; j++)
            {
                Assert.AreEqual(list2.Get(j), j);
                Assert.AreEqual(list2.Get(j + list.Count), j);
            }
            ShortList empty = new ShortList();
            int limit = list.Count;

            for (int j = 0; j < limit; j++)
            {
                Assert.IsTrue(list.AddAll(j, empty));
                Assert.AreEqual(limit, list.Count);
            }
            try
            {
                list.AddAll(limit + 1, empty);
                Assert.Fail("should have thrown an exception");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }

            // try add at beginning
            empty.AddAll(0, list);
            //Assert.AreEqual(empty, list);
            Assert.IsTrue(empty.Equals(list));
            // try in the middle
            empty.AddAll(1, list);
            Assert.AreEqual(2 * list.Count, empty.Count);
            Assert.AreEqual(list.Get(0), empty.Get(0));
            Assert.AreEqual(list.Get(0), empty.Get(1));
            Assert.AreEqual(list.Get(1), empty.Get(2));
            Assert.AreEqual(list.Get(1), empty.Get(6));
            Assert.AreEqual(list.Get(2), empty.Get(3));
            Assert.AreEqual(list.Get(2), empty.Get(7));
            Assert.AreEqual(list.Get(3), empty.Get(4));
            Assert.AreEqual(list.Get(3), empty.Get(8));
            Assert.AreEqual(list.Get(4), empty.Get(5));
            Assert.AreEqual(list.Get(4), empty.Get(9));

            // try at the end
            empty.AddAll(empty.Count, list);
            Assert.AreEqual(3 * list.Count, empty.Count);
            Assert.AreEqual(list.Get(0), empty.Get(0));
            Assert.AreEqual(list.Get(0), empty.Get(1));
            Assert.AreEqual(list.Get(0), empty.Get(10));
            Assert.AreEqual(list.Get(1), empty.Get(2));
            Assert.AreEqual(list.Get(1), empty.Get(6));
            Assert.AreEqual(list.Get(1), empty.Get(11));
            Assert.AreEqual(list.Get(2), empty.Get(3));
            Assert.AreEqual(list.Get(2), empty.Get(7));
            Assert.AreEqual(list.Get(2), empty.Get(12));
            Assert.AreEqual(list.Get(3), empty.Get(4));
            Assert.AreEqual(list.Get(3), empty.Get(8));
            Assert.AreEqual(list.Get(3), empty.Get(13));
            Assert.AreEqual(list.Get(4), empty.Get(5));
            Assert.AreEqual(list.Get(4), empty.Get(9));
            Assert.AreEqual(list.Get(4), empty.Get(14));
        }
        [Test]
        public void TestClear()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 500; j++)
            {
                list.Add(j);
            }
            Assert.AreEqual(500, list.Count);
            list.Clear();
            Assert.AreEqual(0, list.Count);
            for (short j = 0; j < 500; j++)
            {
                list.Add((short)(j + 1));
            }
            Assert.AreEqual(500, list.Count);
            for (short j = 0; j < 500; j++)
            {
                Assert.AreEqual(j + 1, list.Get(j));
            }
        }
        [Test]
        public void TestContains()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j += 2)
            {
                list.Add(j);
            }
            for (short j = 0; j < 1000; j++)
            {
                if (j % 2 == 0)
                {
                    Assert.IsTrue(list.Contains(j));
                }
                else
                {
                    Assert.IsTrue(!list.Contains(j));
                }
            }
        }
        [Test]
        public void TestContainsAll()
        {
            ShortList list = new ShortList();

            Assert.IsTrue(list.ContainsAll(list));
            for (short j = 0; j < 10; j++)
            {
                list.Add(j);
            }
            ShortList list2 = new ShortList(list);

            Assert.IsTrue(list2.ContainsAll(list));
            Assert.IsTrue(list.ContainsAll(list2));
            list2.Add((short)10);
            Assert.IsTrue(list2.ContainsAll(list));
            Assert.IsTrue(!list.ContainsAll(list2));
            list.Add((short)11);
            Assert.IsTrue(!list2.ContainsAll(list));
            Assert.IsTrue(!list.ContainsAll(list2));
        }
        [Test]
        public void TestEquals()
        {
            ShortList list = new ShortList();

            //Assert.AreEqual(list, list);
            Assert.IsTrue(list.Equals(list));
            Assert.IsTrue(!list.Equals(null));
            ShortList list2 = new ShortList(200);

            //Assert.AreEqual(list, list2);
            Assert.IsTrue(list.Equals(list2));
            //Assert.AreEqual(list2, list);
            Assert.IsTrue(list2.Equals(list));
            Assert.AreEqual(list.GetHashCode(), list2.GetHashCode());
            list.Add((short)0);
            list.Add((short)1);
            list2.Add((short)1);
            list2.Add((short)0);
            Assert.IsTrue(!list.Equals(list2));
            list2.RemoveValue((short)1);
            list2.Add((short)1);
            //Assert.AreEqual(list, list2);
            Assert.IsTrue(list.Equals(list2));
            //Assert.AreEqual(list2, list);
            Assert.IsTrue(list2.Equals(list));
            list2.Add((short)2);
            Assert.IsTrue(!list.Equals(list2));
            Assert.IsTrue(!list2.Equals(list));
        }
        [Test]
        public void TestGet()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (short j = 0; j < 1001; j++)
            {
                try
                {
                    Assert.AreEqual(j, list.Get(j));
                    if (j == 1000)
                    {
                        Assert.Fail("should have gotten exception");
                    }
                }
                catch (IndexOutOfRangeException)
                {
                    if (j != 1000)
                    {
                        Assert.Fail("unexpected IndexOutOfRangeException");
                    }
                }
            }
        }
        [Test]
        public void TestIndexOf()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add((short)(j / 2));
            }
            for (short j = 0; j < 1000; j++)
            {
                if (j < 500)
                {
                    Assert.AreEqual(j * 2, list.IndexOf(j));
                }
                else
                {
                    Assert.AreEqual(-1, list.IndexOf(j));
                }
            }
        }
        [Test]
        public void TestIsEmpty()
        {
            ShortList list1 = new ShortList();
            ShortList list2 = new ShortList(1000);
            ShortList list3 = new ShortList(list1);

            Assert.IsTrue(list1.IsEmpty());
            Assert.IsTrue(list2.IsEmpty());
            Assert.IsTrue(list3.IsEmpty());
            list1.Add((short)1);
            list2.Add((short)2);
            list3 = new ShortList(list2);
            Assert.IsTrue(!list1.IsEmpty());
            Assert.IsTrue(!list2.IsEmpty());
            Assert.IsTrue(!list3.IsEmpty());
            list1.Clear();
            list2.Remove(0);
            list3.RemoveValue((short)2);
            Assert.IsTrue(list1.IsEmpty());
            Assert.IsTrue(list2.IsEmpty());
            Assert.IsTrue(list3.IsEmpty());
        }
        [Test]
        public void TestLastIndexOf()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add((short)(j / 2));
            }
            for (short j = 0; j < 1000; j++)
            {
                if (j < 500)
                {
                    Assert.AreEqual(1 + j * 2, list.LastIndexOf(j));
                }
                else
                {
                    Assert.AreEqual(-1, list.IndexOf(j));
                }
            }
        }
        [Test]
        public void TestRemove()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(j, list.Remove(0));
                Assert.AreEqual((short)(999 - j), list.Count);
            }
            for (short j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual((short)(999 - j),
                             list.Remove((short)(999 - j)));
                Assert.AreEqual(999 - j, list.Count);
            }
            try
            {
                list.Remove(0);
                Assert.Fail("should have caught IndexOutOfRangeException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
        }
        [Test]
        public void TestRemoveValue()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add((short)(j / 2));
            }
            for (short j = 0; j < 1000; j++)
            {
                if (j < 500)
                {
                    Assert.IsTrue(list.RemoveValue(j));
                    Assert.IsTrue(list.RemoveValue(j));
                }
                Assert.IsTrue(!list.RemoveValue(j));
            }
        }
        [Test]
        public void TestRemoveAll()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            ShortList listCopy = new ShortList(list);
            ShortList listOdd = new ShortList();
            ShortList listEven = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                if (j % 2 == 0)
                {
                    listEven.Add(j);
                }
                else
                {
                    listOdd.Add(j);
                }
            }
            list.RemoveAll(listEven);
            Assert.IsTrue(list.Equals(listOdd));// Assert.AreEqual(list, listOdd);
            list.RemoveAll(listOdd);
            Assert.IsTrue(list.IsEmpty());
            listCopy.RemoveAll(listOdd);
            //Assert.AreEqual(listCopy, listEven);
            Assert.IsTrue(listCopy.Equals(listEven));
            listCopy.RemoveAll(listEven);
            Assert.IsTrue(listCopy.IsEmpty());
        }
        [Test]
        public void TestRetainAll()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            ShortList listCopy = new ShortList(list);
            ShortList listOdd = new ShortList();
            ShortList listEven = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                if (j % 2 == 0)
                {
                    listEven.Add(j);
                }
                else
                {
                    listOdd.Add(j);
                }
            }
            list.RetainAll(listOdd);
            Assert.IsTrue(list.Equals(listOdd));// Assert.AreEqual(list, listOdd);
            list.RetainAll(listEven);
            Assert.IsTrue(list.IsEmpty());
            listCopy.RetainAll(listEven);
            //Assert.AreEqual(listCopy, listEven);
            Assert.IsTrue(listCopy.Equals(listEven));
            listCopy.RetainAll(listOdd);
            Assert.IsTrue(listCopy.IsEmpty());
        }
        [Test]
        public void TestSet()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (short j = 0; j < 1001; j++)
            {
                try
                {
                    list.Set(j, (short)(j + 1));
                    if (j == 1000)
                    {
                        Assert.Fail("Should have gotten exception");
                    }
                    Assert.AreEqual(j + 1, list.Get(j));
                }
                catch (IndexOutOfRangeException)
                {
                    if (j != 1000)
                    {
                        Assert.Fail("premature exception");
                    }
                }
            }
        }
        [Test]
        public void TestSize()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(j, list.Count);
                list.Add(j);
                Assert.AreEqual(j + 1, list.Count);
            }
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(1000 - j, list.Count);
                list.RemoveValue(j);
                Assert.AreEqual(999 - j, list.Count);
            }
        }
        [Test]
        public void TestToArray()
        {
            ShortList list = new ShortList();

            for (short j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            short[] a1 = list.ToArray();

            Assert.AreEqual(a1.Length, list.Count);
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(a1[j], list.Get(j));
            }
            short[] a2 = new short[list.Count];
            short[] a3 = list.ToArray(a2);

            Assert.AreSame(a2, a3);
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(a2[j], list.Get(j));
            }
            short[] ashort = new short[list.Count - 1];
            short[] aLong = new short[list.Count + 1];
            short[] a4 = list.ToArray(ashort);
            short[] a5 = list.ToArray(aLong);

            Assert.IsTrue(a4 != ashort);
            Assert.IsTrue(a5 != aLong);
            Assert.AreEqual(a4.Length, list.Count);
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(a3[j], list.Get(j));
            }
            Assert.AreEqual(a5.Length, list.Count);
            for (short j = 0; j < 1000; j++)
            {
                Assert.AreEqual(a5[j], list.Get(j));
            }
        }
    }

}