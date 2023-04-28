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

namespace TestCases.SS.Formula.Functions
{

    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula.Eval;
    using TestCases.SS.Formula.Functions;
    using NPOI.SS.Formula.Functions;

    /**
     * Test cases for MATCH()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestMatch
    {
        /** less than or equal to */
        private static NumberEval MATCH_LARGEST_LTE = new NumberEval(1);
        private static NumberEval MATCH_EXACT = new NumberEval(0);
        /** greater than or equal to */
        private static NumberEval MATCH_SMALLEST_GTE = new NumberEval(-1);

        private static StringEval MATCH_INVALID = new StringEval("blabla");

        private static ValueEval invokeMatch(ValueEval Lookup_value, ValueEval Lookup_array, ValueEval match_type)
        {
            ValueEval[] args = { Lookup_value, Lookup_array, match_type, };
            return new Match().Evaluate(args, -1, (short)-1);
        }

        private static ValueEval invokeMatch(ValueEval lookup_value, ValueEval lookup_array)
        {
            ValueEval[] args = { lookup_value, lookup_array, };
            return new Match().Evaluate(args, -1, (short)-1);
        }
        private static void ConfirmInt(int expected, ValueEval actualEval)
        {
            if (!(actualEval is NumericValueEval))
            {
                Assert.Fail("Expected numeric result but had " + actualEval);
            }
            NumericValueEval nve = (NumericValueEval)actualEval;
            Assert.AreEqual(expected, nve.NumberValue, 0);
        }
        [Test]
        public void TestSimpleNumber()
        {

            ValueEval[] values = {
            new NumberEval(4),
            new NumberEval(5),
            new NumberEval(10),
            new NumberEval(10),
            new NumberEval(25),
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A5", values);

            ConfirmInt(2, invokeMatch(new NumberEval(5), ae, MATCH_LARGEST_LTE));
            ConfirmInt(2, invokeMatch(new NumberEval(5), ae));
            ConfirmInt(2, invokeMatch(new NumberEval(5), ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(new NumberEval(10), ae, MATCH_LARGEST_LTE));
            ConfirmInt(3, invokeMatch(new NumberEval(10), ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(new NumberEval(20), ae, MATCH_LARGEST_LTE));
            Assert.AreEqual(ErrorEval.NA, invokeMatch(new NumberEval(20), ae, MATCH_EXACT));
        }
        [Test]
        public void TestReversedNumber()
        {

            ValueEval[] values = {
            new NumberEval(25),
            new NumberEval(10),
            new NumberEval(10),
            new NumberEval(10),
            new NumberEval(4),
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A5", values);

            ConfirmInt(2, invokeMatch(new NumberEval(10), ae, MATCH_SMALLEST_GTE));
            ConfirmInt(2, invokeMatch(new NumberEval(10), ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(new NumberEval(9), ae, MATCH_SMALLEST_GTE));
            ConfirmInt(1, invokeMatch(new NumberEval(20), ae, MATCH_SMALLEST_GTE));
            ConfirmInt(5, invokeMatch(new NumberEval(3), ae, MATCH_SMALLEST_GTE));
            Assert.AreEqual(ErrorEval.NA, invokeMatch(new NumberEval(20), ae, MATCH_EXACT));
            Assert.AreEqual(ErrorEval.NA, invokeMatch(new NumberEval(26), ae, MATCH_SMALLEST_GTE));
        }
        [Test]
        public void TestSimpleString()
        {

            ValueEval[] values = {
            new StringEval("Albert"),
            new StringEval("Charles"),
            new StringEval("Ed"),
            new StringEval("Greg"),
            new StringEval("Ian"),
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A5", values);

            // Note String comparisons are case insensitive
            ConfirmInt(3, invokeMatch(new StringEval("Ed"), ae, MATCH_LARGEST_LTE));
            ConfirmInt(3, invokeMatch(new StringEval("eD"), ae, MATCH_LARGEST_LTE));
            ConfirmInt(3, invokeMatch(new StringEval("Ed"), ae, MATCH_EXACT));
            ConfirmInt(3, invokeMatch(new StringEval("ed"), ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(new StringEval("Hugh"), ae, MATCH_LARGEST_LTE));
            Assert.AreEqual(ErrorEval.NA, invokeMatch(new StringEval("Hugh"), ae, MATCH_EXACT));
        }
        [Test]
        public void TestSimpleWildcardValuesString()
        {
            // Arrange
            ValueEval[] values = {
                new StringEval("Albert"),
                new StringEval("Charles"),
                new StringEval("Ed"),
                new StringEval("Greg"),
                new StringEval("Ian"),
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A5", values);

            // Note String comparisons are case insensitive
            ConfirmInt(3, invokeMatch(new StringEval("e*"), ae, MATCH_EXACT));
            ConfirmInt(3, invokeMatch(new StringEval("*d"), ae, MATCH_EXACT));

            ConfirmInt(1, invokeMatch(new StringEval("Al*"), ae, MATCH_EXACT));
            ConfirmInt(2, invokeMatch(new StringEval("Char*"), ae, MATCH_EXACT));

            ConfirmInt(4, invokeMatch(new StringEval("*eg"), ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(new StringEval("G?eg"), ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(new StringEval("??eg"), ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(new StringEval("G*?eg"), ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(new StringEval("Hugh"), ae, MATCH_LARGEST_LTE));

            ConfirmInt(5, invokeMatch(new StringEval("*Ian*"), ae, MATCH_EXACT));
            ConfirmInt(5, invokeMatch(new StringEval("*Ian*"), ae, MATCH_LARGEST_LTE));
        }
        [Test]
        public void TestTildeString()
        {

            ValueEval[] values = {
                new StringEval("what?"),
                new StringEval("all*")
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A2", values);

            ConfirmInt(1, invokeMatch(new StringEval("what~?"), ae, MATCH_EXACT));
            ConfirmInt(2, invokeMatch(new StringEval("all~*"), ae, MATCH_EXACT));
        }
        [Test]
        public void TestSimpleBoolean()
        {

            ValueEval[] values = {
                BoolEval.FALSE,
                BoolEval.FALSE,
                BoolEval.TRUE,
                BoolEval.TRUE,
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A4", values);

            // Note String comparisons are case insensitive
            ConfirmInt(2, invokeMatch(BoolEval.FALSE, ae, MATCH_LARGEST_LTE));
            ConfirmInt(1, invokeMatch(BoolEval.FALSE, ae, MATCH_EXACT));
            ConfirmInt(4, invokeMatch(BoolEval.TRUE, ae, MATCH_LARGEST_LTE));
            ConfirmInt(3, invokeMatch(BoolEval.TRUE, ae, MATCH_EXACT));
        }
        [Test]
        public void TestHeterogeneous()
        {

            ValueEval[] values = {
                new NumberEval(4),
                BoolEval.FALSE,
                new NumberEval(5),
                new StringEval("Albert"),
                BoolEval.FALSE,
                BoolEval.TRUE,
                new NumberEval(10),
                new StringEval("Charles"),
                new StringEval("Ed"),
                new NumberEval(10),
                new NumberEval(25),
                BoolEval.TRUE,
                new StringEval("Ed"),
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A13", values);

            Assert.AreEqual(ErrorEval.NA, invokeMatch(new StringEval("Aaron"), ae, MATCH_LARGEST_LTE));

            ConfirmInt(5, invokeMatch(BoolEval.FALSE, ae, MATCH_LARGEST_LTE));
            ConfirmInt(2, invokeMatch(BoolEval.FALSE, ae, MATCH_EXACT));
            ConfirmInt(3, invokeMatch(new NumberEval(5), ae, MATCH_LARGEST_LTE));
            ConfirmInt(3, invokeMatch(new NumberEval(5), ae, MATCH_EXACT));

            ConfirmInt(8, invokeMatch(new StringEval("CHARLES"), ae, MATCH_EXACT));

            ConfirmInt(4, invokeMatch(new StringEval("Ben"), ae, MATCH_LARGEST_LTE));

            ConfirmInt(13, invokeMatch(new StringEval("ED"), ae, MATCH_LARGEST_LTE));
            ConfirmInt(9, invokeMatch(new StringEval("ED"), ae, MATCH_EXACT));

            ConfirmInt(13, invokeMatch(new StringEval("Hugh"), ae, MATCH_LARGEST_LTE));
            Assert.AreEqual(ErrorEval.NA, invokeMatch(new StringEval("Hugh"), ae, MATCH_EXACT));

            ConfirmInt(11, invokeMatch(new NumberEval(30), ae, MATCH_LARGEST_LTE));
            ConfirmInt(12, invokeMatch(BoolEval.TRUE, ae, MATCH_LARGEST_LTE));
        }


        /**
         * Ensures that the match_type argument can be an <c>AreaEval</c>.<br/>
         * Bugzilla 44421
         */
        [Test]
        public void TestMatchArgTypeArea()
        {

            ValueEval[] values = {
            new NumberEval(4),
            new NumberEval(5),
            new NumberEval(10),
            new NumberEval(10),
            new NumberEval(25),
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A5", values);

            AreaEval matchAE = EvalFactory.CreateAreaEval("C1:C1", new ValueEval[] { MATCH_LARGEST_LTE, });

            try
            {
                ConfirmInt(4, invokeMatch(new NumberEval(10), ae, matchAE));
            }
            catch (Exception e)
            {
                if (e.Message.StartsWith("Unexpected match_type type"))
                {
                    // identified bug 44421
                    Assert.Fail(e.Message);
                }
                // some other error ??
                throw;
            }
        }

        [Test]
        public void TestInvalidMatchType()
        {

            ValueEval[] values = {
            new NumberEval(4),
            new NumberEval(5),
            new NumberEval(10),
            new NumberEval(10),
            new NumberEval(25),
        };

            AreaEval ae = EvalFactory.CreateAreaEval("A1:A5", values);

            ConfirmInt(2, invokeMatch(new NumberEval(5), ae, MATCH_LARGEST_LTE));

            Assert.AreEqual(ErrorEval.REF_INVALID, invokeMatch(new StringEval("Ben"), ae, MATCH_INVALID), "Should return #REF! for invalid match type");
        }
    }

}