/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.SS.Formula.Functions
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;

    /**
     * Tests for the INDEX() function
     * 
     * @author Josh Micich
     */
    public class TestIndex
    {

        private static double[] TEST_VALUES0 = {
			1, 2,
			3, 4,
			5, 6,
			7, 8,
			9, 10,
			11, 12,
	};

        /**
         * For the case when the first argument to INDEX() is1 an area reference
         */
        public void TestEvaluateAreaReference()
        {

            double[] values = TEST_VALUES0;
            ConfirmAreaEval("C1:D6", values, 4, 1, 7);
            ConfirmAreaEval("C1:D6", values, 6, 2, 12);
            ConfirmAreaEval("C1:D6", values, 3, -1, 5);

            // now treat same data as 3 columns, 4 rows
            ConfirmAreaEval("C10:E13", values, 2, 2, 5);
            ConfirmAreaEval("C10:E13", values, 4, -1, 10);
        }

        /**
         * @param areaRefString in Excel notation e.g. 'D2:E97'
         * @param dValues array of evaluated values for the area reference
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

            double actual = NumericFunctionInvoker.Invoke(new Index(), args);
            Assert.AreEqual(expectedResult, actual, 0D);
        }
    }
}
