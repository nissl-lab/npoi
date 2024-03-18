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

namespace TestCases
{
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;
    using System.Globalization;

    /**
     * A class for testing the POI Junit TestCase utility class
     */
    [TestFixture]
    public class TestPOITestCase
    {
        [Test]
        public void AssertStartsWith()
        {
            POITestCase.AssertStartsWith("Apache POI", "");
            POITestCase.AssertStartsWith("Apache POI", "Apache");
            POITestCase.AssertStartsWith("Apache POI", "Apache POI");
        }

        [Test]
        public void AssertEndsWith()
        {
            POITestCase.AssertEndsWith("Apache POI", "");
            POITestCase.AssertEndsWith("Apache POI", "POI");
            POITestCase.AssertEndsWith("Apache POI", "Apache POI");
        }

        [Test]
        public void AssertContains()
        {
            POITestCase.AssertContains("There is a needle in this haystack", "needle");
            /*try {
                POITestCase.AssertContains("There is gold in this haystack", "needle");
                Assert.Fail("found a needle");
            } catch (final junit.framework.AssertFailedException e) {
                // expected
            }*/
        }

        [Test]
        public void AssertContainsIgnoreCase_Locale()
        {
            POITestCase.AssertContainsIgnoreCase("There is a Needle in this haystack", "needlE", CultureInfo.CurrentCulture);
            // FIXME: test failing case
        }

        [Test]
        public void AssertContainsIgnoreCase()
        {
            POITestCase.AssertContainsIgnoreCase("There is a Needle in this haystack", "needlE");
            // FIXME: test failing case
        }

        [Test]
        public void AssertNotContained()
        {
            POITestCase.AssertNotContained("There is a needle in this haystack", "gold");
            // FIXME: test failing case
        }

        [Test]
        public void AssertMapContains()
        {
            Dictionary<String, String> haystack = new Dictionary<string, string>() { { "needle", "value" } };
            POITestCase.AssertContains(haystack, "needle");
            // FIXME: test failing case
        }


        /**
         * Utility method to Get the value of a private/protected field.
         * Only use this method in test cases!!!
         */
        [Ignore("poi")]
        [Test]
        public void GetFieldValue()
        {
            /*
            Class<? super T> clazz;
            T instance;
            Class<R> fieldType;
            String fieldName;

            R expected;
            R actual = POITestCase.GetFieldValue(clazz, instance, fieldType, fieldName);
            Assert.AreEqual(expected, actual);
            */
        }

        /**
         * Utility method to call a private/protected method.
         * Only use this method in test cases!!!
         */
        [Ignore("poi")]
        [Test]
        public void CallMethod()
        {
            /*
            Class<? super T> clazz;
            T instance;
            Class<R> returnType;
            String methodName;
            Class<?>[] parameterTypes;
            Object[] parameters;

            R expected;
            R actual = POITestCase.CallMethod(clazz, instance, returnType, methodName, parameterTypes, parameters);
            Assert.AreEqual(expected, actual);
            */
        }

        /**
         * Utility method to shallow compare all fields of the objects
         * Only use this method in test cases!!!
         */
        [Ignore("poi")]
        [Test]
        public void AssertReflectEquals()
        {
            /*
            Object expected;
            Object actual;
            POITestCase.AssertReflectEquals(expected, actual);
            */
        }
    }

}