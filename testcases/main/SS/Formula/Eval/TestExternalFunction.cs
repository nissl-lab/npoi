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

using junit.framework.TestCase;

using NPOI.hssf.HSSFTestDataSamples;
using NPOI.SS.Formula.functions.FreeRefFunction;
using NPOI.SS.Formula.udf.DefaultUDFFinder;
using NPOI.SS.Formula.udf.AggregatingUDFFinder;
using NPOI.SS.Formula.udf.UDFFinder;
using NPOI.hssf.UserModel.HSSFCell;
using NPOI.hssf.UserModel.HSSFFormulaEvaluator;
using NPOI.hssf.UserModel.HSSFRow;
using NPOI.hssf.UserModel.HSSFSheet;
using NPOI.hssf.UserModel.HSSFWorkbook;
using NPOI.SS.Formula.OperationEvaluationContext;

/**
 * @author Josh Micich
 * @author Petr Udalau - registering UDFs in workbook and using ToolPacks.
 */
public class TestExternalFunction  {

	private static class MyFunc : FreeRefFunction {
		public MyFunc() {
			//
		}

		public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) {
			if (args.Length != 1 || !(args[0] is StringEval)) {
				return ErrorEval.VALUE_INVALID;
			}
			StringEval input = (StringEval) args[0];
			return new StringEval(input.StringValue + "abc");
		}
	}

	private static class MyFunc2 : FreeRefFunction {
		public MyFunc2() {
			//
		}

		public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) {
			if (args.Length != 1 || !(args[0] is StringEval)) {
				return ErrorEval.VALUE_INVALID;
			}
			StringEval input = (StringEval) args[0];
			return new StringEval(input.StringValue + "abc2");
		}
	}

	/**
	 * Checks that an external function can Get invoked from the formula
	 * Evaluator.
	 */
	public void TestInvoke() {
		HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("testNames.xls");
		HSSFSheet sheet = wb.GetSheetAt(0);

		/**
		 * register the two Test UDFs in a UDF Finder, to be passed to the Evaluator
		 */
		UDFFinder udff1 = new DefaultUDFFinder(new String[] { "myFunc", },
				new FreeRefFunction[] { new MyFunc(), });
		UDFFinder udff2 = new DefaultUDFFinder(new String[] { "myFunc2", },
				new FreeRefFunction[] { new MyFunc2(), });
		UDFFinder udff = new AggregatingUDFFinder(udff1, udff2);


		HSSFRow row = sheet.GetRow(0);
		HSSFCell myFuncCell = row.GetCell(1); // =myFunc("_")

		HSSFCell myFunc2Cell = row.GetCell(2); // =myFunc2("_")

		HSSFFormulaEvaluator fe = HSSFFormulaEvaluator.Create(wb, null, udff);
		Assert.AreEqual("_abc", fe.Evaluate(myFuncCell).StringValue);
		Assert.AreEqual("_abc2", fe.Evaluate(myFunc2Cell).StringValue);
	}
}

