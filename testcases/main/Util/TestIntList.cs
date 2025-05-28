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

    using NUnit.Framework;using NUnit.Framework.Legacy;
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

            ClassicAssert.IsTrue(list.IsEmpty());
            list.Add(0);
            list.Add(1);
            IntList list2 = new IntList(list);

            //ClassicAssert.AreEqual(list, list2);
            ClassicAssert.IsTrue(list.Equals(list2));
            IntList list3 = new IntList(2);

            ClassicAssert.IsTrue(list3.IsEmpty());
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
                ClassicAssert.AreEqual(testArray[j], list.Get(j));
            }
            ClassicAssert.AreEqual(testArray.Length, list.Count);

            // add at the beginning
            list.Add(0, -1);
            ClassicAssert.AreEqual(-1, list.Get(0));
            ClassicAssert.AreEqual(testArray.Length + 1, list.Count);
            for (int j = 0; j < testArray.Length; j++)
            {
                ClassicAssert.AreEqual(testArray[j], list.Get(j + 1));
            }

            // add in the middle
            list.Add(5, 4);
            ClassicAssert.AreEqual(4, list.Get(5));
            ClassicAssert.AreEqual(testArray.Length + 2, list.Count);
            for (int j = 0; j < list.Count; j++)
            {
                ClassicAssert.AreEqual(j - 1, list.Get(j));
            }

            // add at the end
            list.Add(list.Count, 6);
            ClassicAssert.AreEqual(testArray.Length + 3, list.Count);
            for (int j = 0; j < list.Count; j++)
            {
                ClassicAssert.AreEqual(j - 1, list.Get(j));
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
            ClassicAssert.AreEqual(1000, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                ClassicAssert.AreEqual(j, list.Get(j));
            }
            list = new IntList(0);
            for (int j = 0; j < 1000; j++)
            {
                list.Add(0, j);
            }
            ClassicAssert.AreEqual(1000, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                ClassicAssert.AreEqual(j, list.Get(999 - j));
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
            ClassicAssert.AreEqual(2 * list.Count, list2.Count);
            for (int j = 0; j < 5; j++)
            {
                ClassicAssert.AreEqual(list2.Get(j), j);
                ClassicAssert.AreEqual(list2.Get(j + list.Count), j);
            }
            IntList empty = new IntList();
            int limit = list.Count;

            for (int j = 0; j < limit; j++)
            {
                ClassicAssert.IsTrue(list.AddAll(j, empty));
                ClassicAssert.AreEqual(limit, list.Count);
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
            //ClassicAssert.AreEqual(empty, list);
            ClassicAssert.IsTrue(empty.Equals(list));

            // try in the middle
            empty.AddAll(1, list);
            ClassicAssert.AreEqual(2 * list.Count, empty.Count);
            ClassicAssert.AreEqual(list.Get(0), empty.Get(0));
            ClassicAssert.AreEqual(list.Get(0), empty.Get(1));
            ClassicAssert.AreEqual(list.Get(1), empty.Get(2));
            ClassicAssert.AreEqual(list.Get(1), empty.Get(6));
            ClassicAssert.AreEqual(list.Get(2), empty.Get(3));
            ClassicAssert.AreEqual(list.Get(2), empty.Get(7));
            ClassicAssert.AreEqual(list.Get(3), empty.Get(4));
            ClassicAssert.AreEqual(list.Get(3), empty.Get(8));
            ClassicAssert.AreEqual(list.Get(4), empty.Get(5));
            ClassicAssert.AreEqual(list.Get(4), empty.Get(9));

            // try at the end
            empty.AddAll(empty.Count, list);
            ClassicAssert.AreEqual(3 * list.Count, empty.Count);
            ClassicAssert.AreEqual(list.Get(0), empty.Get(0));
            ClassicAssert.AreEqual(list.Get(0), empty.Get(1));
            ClassicAssert.AreEqual(list.Get(0), empty.Get(10));
            ClassicAssert.AreEqual(list.Get(1), empty.Get(2));
            ClassicAssert.AreEqual(list.Get(1), empty.Get(6));
            ClassicAssert.AreEqual(list.Get(1), empty.Get(11));
            ClassicAssert.AreEqual(list.Get(2), empty.Get(3));
            ClassicAssert.AreEqual(list.Get(2), empty.Get(7));
            ClassicAssert.AreEqual(list.Get(2), empty.Get(12));
            ClassicAssert.AreEqual(list.Get(3), empty.Get(4));
            ClassicAssert.AreEqual(list.Get(3), empty.Get(8));
            ClassicAssert.AreEqual(list.Get(3), empty.Get(13));
            ClassicAssert.AreEqual(list.Get(4), empty.Get(5));
            ClassicAssert.AreEqual(list.Get(4), empty.Get(9));
            ClassicAssert.AreEqual(list.Get(4), empty.Get(14));
        }
        [Test]
        public void TestClear()
        {
            IntList list = new IntList();

            for (int j = 0; j < 500; j++)
            {
                list.Add(j);
            }
            ClassicAssert.AreEqual(500, list.Count);
            list.Clear();
            ClassicAssert.AreEqual(0, list.Count);
            for (int j = 0; j < 500; j++)
            {
                list.Add(j + 1);
            }
            ClassicAssert.AreEqual(500, list.Count);
            for (int j = 0; j < 500; j++)
            {
                ClassicAssert.AreEqual(j + 1, list.Get(j));
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
                    ClassicAssert.IsTrue(list.Contains(j));
                }
                else
                {
                    ClassicAssert.IsTrue(!list.Contains(j));
                }
            }
        }
        [Test]
        public void TestContainsAll()
        {
            IntList list = new IntList();

            ClassicAssert.IsTrue(list.ContainsAll(list));
            for (int j = 0; j < 10; j++)
            {
                list.Add(j);
            }
            IntList list2 = new IntList(list);

            ClassicAssert.IsTrue(list2.ContainsAll(list));
            ClassicAssert.IsTrue(list.ContainsAll(list2));
            list2.Add(10);
            ClassicAssert.IsTrue(list2.ContainsAll(list));
            ClassicAssert.IsTrue(!list.ContainsAll(list2));
            list.Add(11);
            ClassicAssert.IsTrue(!list2.ContainsAll(list));
            ClassicAssert.IsTrue(!list.ContainsAll(list2));
        }
        [Test]
        public void TestEquals()
        {
            IntList list = new IntList();

            ClassicAssert.AreEqual(list, list);
            ClassicAssert.IsTrue(!list.Equals(null));
            IntList list2 = new IntList(200);

            ClassicAssert.IsTrue(list.Equals(list2));//ClassicAssert.AreEqual(list, list2);
            ClassicAssert.IsTrue(list2.Equals(list));//ClassicAssert.AreEqual(list2, list);
            ClassicAssert.AreEqual(list.GetHashCode(), list2.GetHashCode());
            list.Add(0);
            list.Add(1);
            list2.Add(1);
            list2.Add(0);
            ClassicAssert.IsTrue(!list.Equals(list2));
            list2.RemoveValue(1);
            list2.Add(1);
            ClassicAssert.IsTrue(list.Equals(list2));//ClassicAssert.AreEqual(list, list2);
            ClassicAssert.IsTrue(list2.Equals(list));//ClassicAssert.AreEqual(list2, list);
            list2.Add(2);
            ClassicAssert.IsTrue(!list.Equals(list2));
            ClassicAssert.IsTrue(!list2.Equals(list));
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
                    ClassicAssert.AreEqual(j, list.Get(j));
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
                    ClassicAssert.AreEqual(j * 2, list.IndexOf(j));
                }
                else
                {
                    ClassicAssert.AreEqual(-1, list.IndexOf(j));
                }
            }
        }
        [Test]
        public void TestIsEmpty()
        {
            IntList list1 = new IntList();
            IntList list2 = new IntList(1000);
            IntList list3 = new IntList(list1);

            ClassicAssert.IsTrue(list1.IsEmpty());
            ClassicAssert.IsTrue(list2.IsEmpty());
            ClassicAssert.IsTrue(list3.IsEmpty());
            list1.Add(1);
            list2.Add(2);
            list3 = new IntList(list2);
            ClassicAssert.IsTrue(!list1.IsEmpty());
            ClassicAssert.IsTrue(!list2.IsEmpty());
            ClassicAssert.IsTrue(!list3.IsEmpty());
            list1.Clear();
            list2.Remove(0);
            list3.RemoveValue(2);
            ClassicAssert.IsTrue(list1.IsEmpty());
            ClassicAssert.IsTrue(list2.IsEmpty());
            ClassicAssert.IsTrue(list3.IsEmpty());
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
                    ClassicAssert.AreEqual(1 + j * 2, list.LastIndexOf(j));
                }
                else
                {
                    ClassicAssert.AreEqual(-1, list.IndexOf(j));
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
                ClassicAssert.AreEqual(j, list.Remove(0));
                ClassicAssert.AreEqual(999 - j, list.Count);
            }
            for (int j = 0; j < 1000; j++)
            {
                list.Add(j);
            }
            for (int j = 0; j < 1000; j++)
            {
                ClassicAssert.AreEqual(999 - j, list.Remove(999 - j));
                ClassicAssert.AreEqual(999 - j, list.Count);
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
                    ClassicAssert.IsTrue(list.RemoveValue(j));
                    ClassicAssert.IsTrue(list.RemoveValue(j));
                }
                ClassicAssert.IsTrue(!list.RemoveValue(j));
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
            //ClassicAssert.AreEqual(list, listOdd);
            ClassicAssert.IsTrue(list.Equals(listOdd));
            list.RemoveAll(listOdd);
            ClassicAssert.IsTrue(list.IsEmpty());
            listCopy.RemoveAll(listOdd);
            //ClassicAssert.AreEqual(listCopy, listEven);
            ClassicAssert.IsTrue(listCopy.Equals(listEven));
            listCopy.RemoveAll(listEven);
            ClassicAssert.IsTrue(listCopy.IsEmpty());
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
            //ClassicAssert.AreEqual(list, listOdd);
            ClassicAssert.IsTrue(list.Equals(listOdd));
            list.RetainAll(listEven);
            ClassicAssert.IsTrue(list.IsEmpty());
            listCopy.RetainAll(listEven);
            //ClassicAssert.AreEqual(listCopy, listEven);
            ClassicAssert.IsTrue(listCopy.Equals(listEven));
            listCopy.RetainAll(listOdd);
            ClassicAssert.IsTrue(listCopy.IsEmpty());
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
                    ClassicAssert.AreEqual(j + 1, list.Get(j));
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
                ClassicAssert.AreEqual(j, list.Count);
                list.Add(j);
                ClassicAssert.AreEqual(j + 1, list.Count);
            }
            for (int j = 0; j < 1000; j++)
            {
                ClassicAssert.AreEqual(1000 - j, list.Count);
                list.RemoveValue(j);
                ClassicAssert.AreEqual(999 - j, list.Count);
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

            ClassicAssert.AreEqual(a1.Length, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                ClassicAssert.AreEqual(a1[j], list.Get(j));
            }
            int[] a2 = new int[list.Count];
            int[] a3 = list.ToArray(a2);

            ClassicAssert.AreSame(a2, a3);
            for (int j = 0; j < 1000; j++)
            {
                ClassicAssert.AreEqual(a2[j], list.Get(j));
            }
            int[] ashort = new int[list.Count - 1];
            int[] aLong = new int[list.Count + 1];
            int[] a4 = list.ToArray(ashort);
            int[] a5 = list.ToArray(aLong);

            ClassicAssert.IsTrue(a4 != ashort);
            ClassicAssert.IsTrue(a5 != aLong);
            ClassicAssert.AreEqual(a4.Length, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                ClassicAssert.AreEqual(a3[j], list.Get(j));
            }
            ClassicAssert.AreEqual(a5.Length, list.Count);
            for (int j = 0; j < 1000; j++)
            {
                ClassicAssert.AreEqual(a5[j], list.Get(j));
            }
        }
    }

}