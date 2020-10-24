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
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Util;
    using NPOI.Util;
    using TestCases.SS.Formula.Functions;
    using NPOI.SS.Formula.Functions;
    using Index = NPOI.SS.Formula.Functions.Index;

    /**
     * Tests for the INDEX() function.</p>
     *
     * This class Contains just a few specific cases that directly invoke {@link Index},
     * with minimum overhead.<br/>
     * Another Test: {@link TestIndexFunctionFromSpreadsheet} operates from a higher level
     * and has far greater coverage of input permutations.<br/>
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestIndex
    {

        private static Index FUNC_INST = new Index();
        private static double[] TEST_VALUES0 = {
			1, 2,
			3, 4,
			5, 6,
			7, 8,
			9, 10,
			11, 12,
	};

        /**
         * For the case when the first argument to INDEX() is an area reference
         */
        [Test]
        public void TestEvaluateAreaReference()
        {

            double[] values = TEST_VALUES0;
            ConfirmAreaEval("C1:D6", values, 4, 1, 7);
            ConfirmAreaEval("C1:D6", values, 6, 2, 12);
            ConfirmAreaEval("C1:D6", values, 3, 1, 5);

            // now treat same data as 3 columns, 4 rows
            ConfirmAreaEval("C10:E13", values, 2, 2, 5);
            ConfirmAreaEval("C10:E13", values, 4, 1, 10);
        }

        /**
         * @param areaRefString in Excel notation e.g. 'D2:E97'
         * @param dValues array of Evaluated values for the area reference
         * @param rowNum 1-based
         * @param colNum 1-based, pass -1 to signify argument not present
         */
        private static void ConfirmAreaEval(String areaRefString, double[] dValues,
                int rowNum, int colNum, double expectedResult)
        {
            ValueEval[] values = new ValueEval[dValues.Length];
            for (int i = 0; i < values.Length; i++)
            {
                values[i] = new NumberEval(dValues[i]);
            }
            AreaEval arg0 = EvalFactory.CreateAreaEval(areaRefString, values);

            ValueEval[] args;
            if (colNum > 0)
            {
                args = new ValueEval[] { arg0, new NumberEval(rowNum), new NumberEval(colNum), };
            }
            else
            {
                args = new ValueEval[] { arg0, new NumberEval(rowNum), };
            }

            double actual = invokeAndDereference(args);
            Assert.AreEqual(expectedResult, actual, 0D);
        }

        private static double invokeAndDereference(ValueEval[] args)
        {
            ValueEval ve = FUNC_INST.Evaluate(args, -1, -1);
            ve = WorkbookEvaluator.DereferenceResult(ve, -1, -1);
            Assert.AreEqual(typeof(NumberEval), ve.GetType());
            return ((NumberEval)ve).NumberValue;
        }

        /**
         * Tests expressions like "INDEX(A1:C1,,2)".<br/>
         * This problem was found while fixing bug 47048 and is observable up to svn r773441.
         */
        [Test]
        public void TestMissingArg()
        {
            ValueEval[] values = {
				new NumberEval(25.0),
				new NumberEval(26.0),
				new NumberEval(28.0),
		};
            AreaEval arg0 = EvalFactory.CreateAreaEval("A10:C10", values);
            ValueEval[] args = new ValueEval[] { arg0, MissingArgEval.instance, new NumberEval(2), };
            ValueEval actualResult;
            try
            {
                actualResult = FUNC_INST.Evaluate(args, -1, -1);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Unexpected arg eval type (NPOI.hssf.Record.Formula.Eval.MissingArgEval"))
                {
                    throw new AssertionException("Identified bug 47048b - INDEX() should support missing-arg");
                }
                throw e;
            }
            // result should be an area eval "B10:B10"
            AreaEval ae = ConfirmAreaEval("B10:B10", actualResult);
            actualResult = ae.GetValue(0, 0);
            Assert.AreEqual(typeof(NumberEval), actualResult.GetType());
            Assert.AreEqual(26.0, ((NumberEval)actualResult).NumberValue, 0.0);
        }

        /**
         * When the argument to INDEX is a reference, the result should be a reference
         * A formula like "OFFSET(INDEX(A1:B2,2,1),1,1,1,1)" should return the value of cell B3.
         * This works because the INDEX() function returns a reference to A2 (not the value of A2)
         */
        [Test]
        public void TestReferenceResult()
        {
            ValueEval[] values = new ValueEval[4];
            Arrays.Fill(values, NumberEval.ZERO);
            AreaEval arg0 = EvalFactory.CreateAreaEval("A1:B2", values);
            ValueEval[] args = new ValueEval[] { arg0, new NumberEval(2), new NumberEval(1), };
            ValueEval ve = FUNC_INST.Evaluate(args, -1, -1);
            ConfirmAreaEval("A2:A2", ve);
        }

        /**
         * Confirms that the result is an area ref with the specified coordinates
         * @return <c>ve</c> cast to {@link AreaEval} if it is valid
         */
        private static AreaEval ConfirmAreaEval(String refText, ValueEval ve)
        {
            CellRangeAddress cra = CellRangeAddress.ValueOf(refText);
            Assert.IsTrue(ve is AreaEval);
            AreaEval ae = (AreaEval)ve;
            Assert.AreEqual(cra.FirstRow, ae.FirstRow);
            Assert.AreEqual(cra.FirstColumn, ae.FirstColumn);
            Assert.AreEqual(cra.LastRow, ae.LastRow);
            Assert.AreEqual(cra.LastColumn, ae.LastColumn);
            return ae;
        }
    }

}