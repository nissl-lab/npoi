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

namespace NPOI.SS.Formula.functions;

using junit.framework.Test;
using junit.framework.TestSuite;

/**
 * Direct Tests for all implementors of <code>Function</code>.
 *
 * @author Josh Micich
 */
public class AllIndividualFunctionEvaluationTests {

	public static Test suite() {
		TestSuite result = new TestSuite(AllIndividualFunctionEvaluationTests.class.GetName());
		result.AddTestSuite(TestAverage.class);
		result.AddTestSuite(TestCountFuncs.class);
		result.AddTestSuite(TestDate.class);
		result.AddTestSuite(TestDays360.class);
		result.AddTestSuite(TestFinanceLib.class);
		result.AddTestSuite(TestFind.class);
		result.AddTestSuite(TestIndex.class);
		result.AddTestSuite(TestIndexFunctionFromSpreadsheet.class);
		result.AddTestSuite(TestIndirect.class);
		result.AddTestSuite(TestIsBlank.class);
		result.AddTestSuite(TestLen.class);
		result.AddTestSuite(TestLookupFunctionsFromSpreadsheet.class);
		result.AddTestSuite(TestMatch.class);
		result.AddTestSuite(TestMathX.class);
		result.AddTestSuite(TestMid.class);
		result.AddTestSuite(TestNper.class);
		result.AddTestSuite(TestOffset.class);
		result.AddTestSuite(TestPmt.class);
		result.AddTestSuite(TestRoundFuncs.class);
		result.AddTestSuite(TestRowCol.class);
		result.AddTestSuite(TestStatsLib.class);
		result.AddTestSuite(TestSubtotal.class);
		result.AddTestSuite(TestSumif.class);
		result.AddTestSuite(TestSumproduct.class);
		result.AddTestSuite(TestText.class);
		result.AddTestSuite(TestTFunc.class);
		result.AddTestSuite(TestTime.class);
		result.AddTestSuite(TestTrim.class);
		result.AddTestSuite(TestTrunc.class);
		result.AddTestSuite(TestValue.class);
		result.AddTestSuite(TestXYNumericFunction.class);
		result.AddTestSuite(TestAddress.class);
		result.AddTestSuite(TestClean.class);
		return result;
	}
}

