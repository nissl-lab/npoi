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

using System;
using System.Text;
using NUnit.Framework;
using System.Collections.Generic;
using NPOI.Util;
using System.Reflection;
using System.Globalization;

namespace TestCases
{

    /**
     * Parent class for POI JUnit TestCases, which provide additional
     *  features 
     */
    public class POITestCase
    {
        public static void AssertContains(String haystack, String needle)
        {
            Assert.IsTrue(
                  haystack.Contains(needle),
                  "Unable to find expected text '" + needle + "' in text:\n" + haystack
            );
        }
        public static void AssertContainsIgnoreCase(String haystack, String needle, CultureInfo locale)
        {
            Assert.IsNotNull(haystack);
            Assert.IsNotNull(needle);
            String hay = haystack.ToLower(locale);
            String n = needle.ToLower(locale);
            Assert.IsTrue(hay.Contains(n), "Unable to find expected text '" + needle + "' in1 text:\n" + haystack);
        }
        public static void AssertContainsIgnoreCase(String haystack, String needle)
        {
            AssertContainsIgnoreCase(haystack, needle, CultureInfo.CurrentCulture);
        }

        public static void AssertNotContained(String haystack, String needle)
        {
            Assert.IsFalse(haystack.Contains(needle),
                  "Unexpectedly found text '" + needle + "' in text:\n" + haystack
            );
        }

        /**
         * @param map haystack
         * @param key needle
         */
        public static void AssertContains<TKey, TValue>(Dictionary<TKey, TValue> map, TKey key)
        {
            if (map.ContainsKey(key))
            {
                return;
            }
            Assert.Fail("Unable to find " + key + " in " + map);
        }
        public static void AssertEquals<T>(T[] expected, T[] actual)
        {
            Assert.AreEqual(expected.Length, actual.Length, "Non-matching lengths");
            for (int i = 0; i < expected.Length; i++)
            {
                Assert.AreEqual(expected[i], actual[i], "Mis-match at offset " + i);
            }
        }
        public static void AssertEquals(byte[] expected, byte[] actual)
        {
            Assert.AreEqual(expected.Length, actual.Length, "Non-matching lengths");
            for (int i = 0; i < expected.Length; i++)
            {
                Assert.AreEqual(expected[i], actual[i], "Mis-match at offset " + i);
            }
        }
        public static void AssertContains<T>(T needle, T[] haystack)
        {
            // Check
            foreach (T thing in haystack)
            {
                if (thing.Equals(needle))
                {
                    return;
                }
            }

            // Failed, try to build a nice error
            StringBuilder sb = new StringBuilder();
            sb.Append("Unable to find ").Append(needle).Append(" in [");
            foreach (T thing in haystack)
            {
                sb.Append(" ").Append(thing.ToString()).Append(" ,");
            }
            sb[sb.Length - 1] = ']';

            Assert.Fail(sb.ToString());
        }

        public static void AssertContains<T>(T needle, IList<T> haystack)
        {
            if (haystack.Contains(needle))
            {
                return;
            }
            Assert.Fail("Unable to find " + needle + " in " + haystack);
        }

        public static R GetFieldValue<R, T>(Type clazz, T instance, Type fieldType, String fieldName)
        {
            Assert.IsTrue(clazz.FullName.StartsWith("NPOI"), "Reflection of private fields is only allowed for POI classes.");
            try
            {
                FieldInfo fieldInfo = clazz.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.GetField | BindingFlags.Instance);
                return (R)fieldInfo.GetValue(instance);

            }
            catch (Exception pae)
            {
                throw new RuntimeException("Cannot access field '" + fieldName + "' of class " + clazz, pae.InnerException);
            }
        }

        public static void SkipTest(Exception e)
        {
            Assume.That(false, "This test currently fails with " + e);
        }

        public static void TestPassesNow(int bug)
        {
            Assert.Fail("This test passes now. Please update the unit test and bug " + bug + ".");
        }

        public static void AssertBetween(String message, int value, int min, int max)
        {
            Assert.IsTrue(min <= value, message + ": " + value + " is less than the minimum value of " + min);
            Assert.IsTrue(value <= max, message + ": " + value + " is greater than the maximum value of " + max);
        }

        public static void AssertStrictlyBetween(String message, int value, int min, int max)
        {
            Assert.IsTrue(min < value, message + ": " + value + " is less than or equal to the minimum value of " + min);
            Assert.IsTrue(value < max, message + ": " + value + " is greater than or equal to the maximum value of " + max);
        }
    }
}