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
     * Class to Test IntList
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestIntList
    {
        [Test]
        public void TestConstructors()
        {
            IntList list = new IntList();

            Assert.IsTrue(list.IsEmpty());
            list.Add(0);
            list.Add(1);
            IntList list2 = new IntList(list);

            //Assert.AreEqual(list, list2);
            Assert.IsTrue(list.Equals(list2));
            IntList list3 = new IntList(2);

            Assert.IsTrue(list3.IsEmpty());
        }
        [Test]
        public void TestAdd()
        {
            IntList list = new IntList();
            int[] testArray =
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
            list.Add(0, -1);
            Assert.AreEqual(-1, list.Get(0));
            Assert.AreEqual(testArray.Length + 1, list.Count);
            for (int j = 0; j < testArray.Length; j++)
            {
                Assert.AreEqual(testArray[j], list.Get(j + 1));
            }

            // add in the middle
            list.Add(5, 4);
            Assert.AreEqual(4, list.Get(5));
            Assert.AreEqual(testArray.Length + 2, list.Count);
            for (int j = 0; j < list.Count; j++)
            {
                Assert.AreEqual(j - 1, list.Get(j));
            }

            // add at the end
            list.Add(list.Count, 6);
            Assert.AreEqual(testArray.Length + 3, list.Count);
            for (int j = 0; j < list.Count; j++)
            {
                Assert.AreEqual(j - 1, list.Get(j));
            }

            // add past end
            try
            {
                list.Add(list.Count + 1, 8);
                Assert.Fail("should have thrown exception");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }

            // Test growth
            list = new IntList(0);
            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            Assert.AreEqual(1000, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(j, list.Get(j));
            }
            list = new IntList(0);
            for (int j = 0; j < 1000; j++)
            {
                list.Add(0, j);
            }
            Assert.AreEqual(1000, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(j, list.Get(999 - j));
            }
        }
        [Test]
        public void TestAddAll()
        {
            IntList list = new IntList();

            for (int j = 0; j < 5; j++)
            {
                list.Add(j);
            }
            IntList list2 = new IntList(0);

            list2.AddAll(list);
            list2.AddAll(list);
            Assert.AreEqual(2 * list.Count, list2.Count);
            for (int j = 0; j < 5; j++)
            {
                Assert.AreEqual(list2.Get(j), j);
                Assert.AreEqual(list2.Get(j + list.Count), j);
            }
            IntList empty = new IntList();
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
            IntList list = new IntList();

            for (int j = 0; j < 500; j++)
            {
                list.Add(j);
            }
            Assert.AreEqual(500, list.Count);
            list.Clear();
            Assert.AreEqual(0, list.Count);
            for (int j = 0; j < 500; j++)
            {
                list.Add(j + 1);
            }
            Assert.AreEqual(500, list.Count);
            for (int j = 0; j < 500; j++)
            {
                Assert.AreEqual(j + 1, list.Get(j));
            }
        }
        [Test]
        public void TestContains()
        {
            IntList list = new IntList();

            for (int j = 0; j < 1000; j += 2)
            {
                list.Add(j);
            }
            for (int j = 0; j < 1000; j++)
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
            IntList list = new IntList();

            Assert.IsTrue(list.ContainsAll(list));
            for (int j = 0; j < 10; j++)
            {
                list.Add(j);
            }
            IntList list2 = new IntList(list);

            Assert.IsTrue(list2.ContainsAll(list));
            Assert.IsTrue(list.ContainsAll(list2));
            list2.Add(10);
            Assert.IsTrue(list2.ContainsAll(list));
            Assert.IsTrue(!list.ContainsAll(list2));
            list.Add(11);
            Assert.IsTrue(!list2.ContainsAll(list));
            Assert.IsTrue(!list.ContainsAll(list2));
        }
        [Test]
        public void TestEquals()
        {
            IntList list = new IntList();

            Assert.AreEqual(list, list);
            Assert.IsTrue(!list.Equals(null));
            IntList list2 = new IntList(200);

            Assert.IsTrue(list.Equals(list2));//Assert.AreEqual(list, list2);
            Assert.IsTrue(list2.Equals(list));//Assert.AreEqual(list2, list);
            Assert.AreEqual(list.GetHashCode(), list2.GetHashCode());
            list.Add(0);
            list.Add(1);
            list2.Add(1);
            list2.Add(0);
            Assert.IsTrue(!list.Equals(list2));
            list2.RemoveValue(1);
            list2.Add(1);
            Assert.IsTrue(list.Equals(list2));//Assert.AreEqual(list, list2);
            Assert.IsTrue(list2.Equals(list));//Assert.AreEqual(list2, list);
            list2.Add(2);
            Assert.IsTrue(!list.Equals(list2));
            Assert.IsTrue(!list2.Equals(list));
        }
        [Test]
        public void TestGet()
        {
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (int j = 0; j < 1001; j++)
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
                        Assert.Fail("unexpected IndexOutOfBoundsException");
                    }
                }
            }
        }
        [Test]
        public void TestIndexOf()
        {
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j / 2);
            }
            for (int j = 0; j < 1000; j++)
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
            IntList list1 = new IntList();
            IntList list2 = new IntList(1000);
            IntList list3 = new IntList(list1);

            Assert.IsTrue(list1.IsEmpty());
            Assert.IsTrue(list2.IsEmpty());
            Assert.IsTrue(list3.IsEmpty());
            list1.Add(1);
            list2.Add(2);
            list3 = new IntList(list2);
            Assert.IsTrue(!list1.IsEmpty());
            Assert.IsTrue(!list2.IsEmpty());
            Assert.IsTrue(!list3.IsEmpty());
            list1.Clear();
            list2.Remove(0);
            list3.RemoveValue(2);
            Assert.IsTrue(list1.IsEmpty());
            Assert.IsTrue(list2.IsEmpty());
            Assert.IsTrue(list3.IsEmpty());
        }
        [Test]
        public void TestLastIndexOf()
        {
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j / 2);
            }
            for (int j = 0; j < 1000; j++)
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
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(j, list.Remove(0));
                Assert.AreEqual(999 - j, list.Count);
            }
            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(999 - j, list.Remove(999 - j));
                Assert.AreEqual(999 - j, list.Count);
            }
            try
            {
                list.Remove(0);
                Assert.Fail("should have caught IndexOutOfBoundsException");
            }
            catch (IndexOutOfRangeException)
            {

                // as expected
            }
        }
        [Test]
        public void TestRemoveValue()
        {
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j / 2);
            }
            for (int j = 0; j < 1000; j++)
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
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            IntList listCopy = new IntList(list);
            IntList listOdd = new IntList();
            IntList listEven = new IntList();

            for (int j = 0; j < 1000; j++)
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
            //Assert.AreEqual(list, listOdd);
            Assert.IsTrue(list.Equals(listOdd));
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
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            IntList listCopy = new IntList(list);
            IntList listOdd = new IntList();
            IntList listEven = new IntList();

            for (int j = 0; j < 1000; j++)
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
            //Assert.AreEqual(list, listOdd);
            Assert.IsTrue(list.Equals(listOdd));
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
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (int j = 0; j < 1001; j++)
            {
                try
                {
                    list.Set(j, j + 1);
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
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(j, list.Count);
                list.Add(j);
                Assert.AreEqual(j + 1, list.Count);
            }
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(1000 - j, list.Count);
                list.RemoveValue(j);
                Assert.AreEqual(999 - j, list.Count);
            }
        }
        [Test]
        public void TestToArray()
        {
            IntList list = new IntList();

            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            int[] a1 = list.ToArray();

            Assert.AreEqual(a1.Length, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(a1[j], list.Get(j));
            }
            int[] a2 = new int[list.Count];
            int[] a3 = list.ToArray(a2);

            Assert.AreSame(a2, a3);
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(a2[j], list.Get(j));
            }
            int[] ashort = new int[list.Count - 1];
            int[] aLong = new int[list.Count + 1];
            int[] a4 = list.ToArray(ashort);
            int[] a5 = list.ToArray(aLong);

            Assert.IsTrue(a4 != ashort);
            Assert.IsTrue(a5 != aLong);
            Assert.AreEqual(a4.Length, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(a3[j], list.Get(j));
            }
            Assert.AreEqual(a5.Length, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                Assert.AreEqual(a5[j], list.Get(j));
            }
        }
    }

}