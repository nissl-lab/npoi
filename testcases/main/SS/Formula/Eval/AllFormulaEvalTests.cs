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

namespace NPOI.SS.Formula.Eval;

using junit.framework.Test;
using junit.framework.TestSuite;

/**
 * Collects all Tests the namespace <tt>NPOI.hssf.Record.Formula.Eval</tt>.
 *
 * @author Josh Micich
 */
public class AllFormulaEvalTests {

	public static Test suite() {
		TestSuite result = new TestSuite(AllFormulaEvalTests.class.GetName());
		result.AddTestSuite(TestAreaEval.class);
		result.AddTestSuite(TestCircularReferences.class);
		result.AddTestSuite(TestDivideEval.class);
		result.AddTestSuite(TestEqualEval.class);
		result.AddTestSuite(TestExternalFunction.class);
		result.AddTestSuite(TestFormulaBugs.class);
		result.AddTestSuite(TestFormulasFromSpreadsheet.class);
		result.AddTestSuite(TestMinusZeroResult.class);
		result.AddTestSuite(TestMissingArgEval.class);
		result.AddTestSuite(TestPercentEval.class);
		result.AddTestSuite(TestRangeEval.class);
		result.AddTestSuite(TestUnaryPlusEval.class);
		return result;
	}
}

