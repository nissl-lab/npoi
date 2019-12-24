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

    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NUnit.Framework;

    [TestFixture]
    public class TestWeekdayFunc
    {
        [Test]
        public void TestEvaluate()
        {
            Assert.AreEqual(2.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(2.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(1.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(1.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(2.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(0.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(3.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(1.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(11.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(7.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(12.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(6.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(13.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(5.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(14.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(4.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(15.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(3.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(16.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(2.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(17.0) }, 0, 0)).NumberValue, 0.001);

            Assert.AreEqual(3.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(3.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(1.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(2.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(2.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(1.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(3.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(2.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(11.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(1.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(12.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(7.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(13.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(6.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(14.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(5.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(15.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(4.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(16.0) }, 0, 0)).NumberValue, 0.001);
            Assert.AreEqual(3.0, ((NumberEval)WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(39448.0), new NumberEval(17.0) }, 0, 0)).NumberValue, 0.001);
        }

        [Test]
        public void TestEvaluateInvalid()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, WeekdayFunc.instance.Evaluate(new ValueEval[] { }, 0, 0));
            Assert.AreEqual(ErrorEval.VALUE_INVALID, WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(1.0), new NumberEval(1.0) }, 0, 0));

            Assert.AreEqual(ErrorEval.NUM_ERROR, WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(-1.0) }, 0, 0));
            Assert.AreEqual(ErrorEval.VALUE_INVALID, WeekdayFunc.instance.Evaluate(new ValueEval[] { new StringEval("") }, 0, 0));
            Assert.AreEqual(ErrorEval.VALUE_INVALID, WeekdayFunc.instance.Evaluate(new ValueEval[] { new StringEval("1"), new StringEval("") }, 0, 0));
            Assert.AreEqual(ErrorEval.NUM_ERROR, WeekdayFunc.instance.Evaluate(new ValueEval[] { new StringEval("2"), BlankEval.instance }, 0, 0));
            Assert.AreEqual(ErrorEval.NUM_ERROR, WeekdayFunc.instance.Evaluate(new ValueEval[] { new StringEval("3"), MissingArgEval.instance }, 0, 0));
            Assert.AreEqual(ErrorEval.NUM_ERROR, WeekdayFunc.instance.Evaluate(new ValueEval[] { new NumberEval(1.0), new NumberEval(18.0) }, 0, 0));
        }
    }
}