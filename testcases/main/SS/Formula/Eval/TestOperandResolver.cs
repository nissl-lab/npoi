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

using System;
using NUnit.Framework;
using NPOI.SS.Formula.Eval;
namespace TestCases.SS.Formula.Eval
{

    /**
     * Tests for <c>OperandResolver</c>
     *
     * @author Brendan Nolan
     */
    [TestFixture]
    public class TestOperandResolver
    {
        [Test]
        public void TestParseDouble_bug48472()
        {

            String value = "-";

            Double ResolvedValue;

            try
            {
                ResolvedValue = OperandResolver.ParseDouble(value);
            }
            catch (Exception)
            {
                throw new AssertionException("Identified bug 48472");
            }

            Assert.AreEqual(double.NaN, ResolvedValue);

        }
        [Test]
        public void TestParseDouble_bug49723()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            String value = ".1";

            Double ResolvedValue;

            ResolvedValue = OperandResolver.ParseDouble(value);

            Assert.IsNotNull(ResolvedValue, "Identified bug 49723");

        }

        /**
         * 
         * Tests that a list of valid strings all return a non null value from {@link OperandResolver#ParseDouble(String)}
         * 
         */
        [Test]
        public void TestParseDoubleValidStrings()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            String[] values = new String[] { ".19", "0.19", "1.9", "1E4", "-.19", "-0.19", "8.5", "-1E4", ".5E6", "+1.5", "+1E5", "  +1E5  " };

            foreach (String value in values)
            {
                Assert.AreNotEqual(double.NaN,OperandResolver.ParseDouble(value));  //this bug is caused by double.Parse
                Assert.AreEqual(OperandResolver.ParseDouble(value), Double.Parse(value));
            }

        }

        /**
         * 
         * Tests that a list of invalid strings all return null from {@link OperandResolver#ParseDouble(String)}
         * 
         */
        [Test]
        public void TestParseDoubleInvalidStrings()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            String[] values = new String[] { "-", "ABC", "-X", "1E5a", "Infinity", "NaN", ".5F" };    //, "1,000" };

            foreach (String value in values)
            {
                Assert.AreEqual(double.NaN, OperandResolver.ParseDouble(value));
            }

        }

    }

}