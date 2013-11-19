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

namespace TestCases.SS.Formula.Atp
{
    using System;
    using NPOI.SS.Formula.Atp;
    using NPOI.SS.Formula.Eval;
    using NUnit.Framework;

    /**
     * @author jfaenomoto@gmail.com
     */
    [TestFixture]
    public class TestDateParser
    {
        [Test]
        public void TestFailWhenNoDate()
        {
            try
            {
                DateParser.ParseDate("potato");
                Assert.Fail("Shouldn't parse potato!");
            }
            catch (EvaluationException e)
            {
                Assert.AreEqual(ErrorEval.VALUE_INVALID, e.GetErrorEval());
            }
        }

        [Test]
        public void TestFailWhenLooksLikeDateButItIsnt()
        {
            try
            {
                DateParser.ParseDate("potato/cucumber/banana");
                Assert.Fail("Shouldn't parse this thing!");
            }
            catch (EvaluationException e)
            {
                Assert.AreEqual(ErrorEval.VALUE_INVALID, e.GetErrorEval());
            }
        }

        [Test]
        public void TestFailWhenIsInvalidDate()
        {
            try
            {
                DateParser.ParseDate("13/13/13");
                Assert.Fail("Shouldn't parse this thing!");
            }
            catch (EvaluationException e)
            {
                Assert.AreEqual(ErrorEval.VALUE_INVALID, e.GetErrorEval());
            }
        }

        [Test]
        public void TestShouldParseValidDate()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            Assert.AreEqual(new DateTime(1984, 10, 20), DateParser.ParseDate("1984/10/20"));
        }

        [Test]
        public void TestShouldIgnoreTimestamp()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            Assert.AreEqual(new DateTime(1984, 10, 20), DateParser.ParseDate("1984/10/20 12:34:56"));
        }
    }
}