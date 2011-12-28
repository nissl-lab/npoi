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






using junit.framework.TestCase;
using NPOI.SS.Formula.Eval.ErrorEval;
using NPOI.SS.Formula.Eval.NumberEval;
using NPOI.SS.Formula.Eval.ValueEval;
using NPOI.SS.Formula.Eval.StringEval;

/**
 * Test case for TEXT()
 *
 * @author Stephen Wolke (smwolke at geistig.com)
 */
public class TestText  {
	private static TextFunction T = null;

	public void TestTextWithStringFirstArg() {

		ValueEval strArg = new StringEval("abc");
		ValueEval formatArg = new StringEval("abc");
		ValueEval[] args = { strArg, formatArg };
		ValueEval result = T.TEXT.Evaluate(args, -1, (short)-1);
		Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
	}

	public void TestTextWithDeciamlFormatSecondArg() {

		ValueEval numArg = new NumberEval(321321.321);
		ValueEval formatArg = new StringEval("#,###.00000");
		ValueEval[] args = { numArg, formatArg };
		ValueEval result = T.TEXT.Evaluate(args, -1, (short)-1);
		char groupSeparator = new DecimalFormatSymbols(Locale.GetDefault()).GetGroupingSeparator();
		char decimalSeparator = new DecimalFormatSymbols(Locale.GetDefault()).GetDecimalSeparator();
		ValueEval TestResult = new StringEval("321" + groupSeparator + "321" + decimalSeparator + "32100");
		Assert.AreEqual(testResult.ToString(), result.ToString());
		numArg = new NumberEval(321.321);
		formatArg = new StringEval("00000.00000");
		args[0] = numArg;
		args[1] = formatArg;
		result = T.TEXT.Evaluate(args, -1, (short)-1);
	 TestResult = new StringEval("00321" + decimalSeparator + "32100");
		Assert.AreEqual(testResult.ToString(), result.ToString());

		formatArg = new StringEval("$#.#");
		args[1] = formatArg;
		result = T.TEXT.Evaluate(args, -1, (short)-1);
	 TestResult = new StringEval("$321" + decimalSeparator + "3");
		Assert.AreEqual(testResult.ToString(), result.ToString());
	}

	public void TestTextWithFractionFormatSecondArg() {

		ValueEval numArg = new NumberEval(321.321);
		ValueEval formatArg = new StringEval("# #/#");
		ValueEval[] args = { numArg, formatArg };
		ValueEval result = T.TEXT.Evaluate(args, -1, (short)-1);
		ValueEval TestResult = new StringEval("321 1/3");
		Assert.AreEqual(testResult.ToString(), result.ToString());

		formatArg = new StringEval("# #/##");
		args[1] = formatArg;
		result = T.TEXT.Evaluate(args, -1, (short)-1);
	 TestResult = new StringEval("321 26/81");
		Assert.AreEqual(testResult.ToString(), result.ToString());

		formatArg = new StringEval("#/##");
		args[1] = formatArg;
		result = T.TEXT.Evaluate(args, -1, (short)-1);
	 TestResult = new StringEval("26027/81");
		Assert.AreEqual(testResult.ToString(), result.ToString());
	}

	public void TestTextWithDateFormatSecondArg() {

		ValueEval numArg = new NumberEval(321.321);
		ValueEval formatArg = new StringEval("dd:MM:yyyy hh:mm:ss");
		ValueEval[] args = { numArg, formatArg };
		ValueEval result = T.TEXT.Evaluate(args, -1, (short)-1);
		ValueEval TestResult = new StringEval("16:11:1900 07:42:14");
		Assert.AreEqual(testResult.ToString(), result.ToString());

		// this line is intended to compute how "November" would look like in the current locale
		String november = new SimpleDateFormat("MMMM").format(new GregorianCalendar(2010,10,15).GetTime());

		formatArg = new StringEval("MMMM dd, yyyy");
		args[1] = formatArg;
		result = T.TEXT.Evaluate(args, -1, (short)-1);
	 TestResult = new StringEval(november + " 16, 1900");
		Assert.AreEqual(testResult.ToString(), result.ToString());
	}
}

