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

namespace TestCases.SS.Formula.Eval
{
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.Functions;
    using TestCases.SS.Formula.Functions;
    using NUnit.Framework;
    using NPOI.SS.Formula.Eval;

    /**
     * Tests for <c>AreaEval</c>
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestAreaEval
    {
        [Test]
        public void TestGetValue_bug44950()
        {
            // TODO - this Test probably isn't Testing much anymore
            AreaPtg ptg = new AreaPtg("B2:D3");
            NumberEval one = new NumberEval(1);
            ValueEval[] values = {
				one,
				new NumberEval(2),
				new NumberEval(3),
				new NumberEval(4),
				new NumberEval(5),
				new NumberEval(6),
		};
            AreaEval ae = EvalFactory.CreateAreaEval(ptg, values);
            if (one == ae.GetAbsoluteValue(1, 2))
            {
                throw new AssertionException("Identified bug 44950 a");
            }
            Confirm(1, ae, 1, 1);
            Confirm(2, ae, 1, 2);
            Confirm(3, ae, 1, 3);
            Confirm(4, ae, 2, 1);
            Confirm(5, ae, 2, 2);
            Confirm(6, ae, 2, 3);

        }

        private static void Confirm(int expectedValue, AreaEval ae, int row, int col)
        {
            NumberEval v = (NumberEval)ae.GetAbsoluteValue(row, col);
            Assert.AreEqual(expectedValue, v.NumberValue, 0.0);
        }
    }

}