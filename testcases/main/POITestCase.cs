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
    }
}