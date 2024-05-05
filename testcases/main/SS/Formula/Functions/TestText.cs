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

    using NPOI.SS.Formula.Eval;
    using NUnit.Framework;
    using System;
    using NPOI.SS.Util;
    using NPOI.SS.Formula.Functions;
    using System.Globalization;

    /**
     * Test case for TEXT()
     *
     * @author Stephen Wolke (smwolke at geistig.com)
     */
    [TestFixture]
    public class TestText
    {
        //private static TextFunction T = null;
        [Test]
        public void TestTextWithStringFirstArg()
        {

            ValueEval strArg = new StringEval("abc");
            ValueEval formatArg = new StringEval("abc");
            ValueEval[] args = { strArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            Assert.AreEqual(strArg, result);
        }
        [Test]
        public void TestTextWithDeciamlFormatSecondArg()
        {
            ValueEval numArg = new NumberEval(321321.321);
            ValueEval formatArg = new StringEval("#,###.00000");
            ValueEval[] args = { numArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            //char groupSeparator = new DecimalFormatSymbols(Locale.GetDefault()).GetGroupingSeparator();
            //char decimalSeparator = new DecimalFormatSymbols(Locale.GetDefault()).GetDecimalSeparator();
            
            NumberFormatInfo fs = CultureInfo.GetCultureInfo("en-US").NumberFormat;
            string groupSeparator = fs.NumberGroupSeparator;
            string decimalSeparator = fs.NumberDecimalSeparator; ;

            ValueEval testResult = new StringEval("321" + groupSeparator + "321" + decimalSeparator + "32100");
            Assert.AreEqual(testResult.ToString(), result.ToString());
            numArg = new NumberEval(321.321);
            formatArg = new StringEval("00000.00000");
            args[0] = numArg;
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            testResult = new StringEval("00321" + decimalSeparator + "32100");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            formatArg = new StringEval("$#.#");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            testResult = new StringEval("$321" + decimalSeparator + "3");
            Assert.AreEqual(testResult.ToString(), result.ToString());
        }
        [Test]
        public void TestTextWithFractionFormatSecondArg()
        {

            ValueEval numArg = new NumberEval(321.321);
            ValueEval formatArg = new StringEval("# #/#");
            ValueEval[] args = { numArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            ValueEval testResult = new StringEval("321 1/3");
            Assert.AreEqual(testResult.ToString(), result.ToString());  //this bug is caused by DecimalFormat

            formatArg = new StringEval("# #/##");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            testResult = new StringEval("321 26/81");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            formatArg = new StringEval("#/##");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            testResult = new StringEval("26027/81");
            Assert.AreEqual(testResult.ToString(), result.ToString());
        }
        [Test]
        public void TestTextWithDateFormatSecondArg()
        {
            // Test with Java style M=Month
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("en-US");
            ValueEval numArg = new NumberEval(321.321);
            ValueEval formatArg = new StringEval("dd:MM:yyyy hh:mm:ss");
            ValueEval[] args = { numArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            ValueEval testResult = new StringEval("16:11:1900 07:42:14");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            // Excel also supports "m before h is month"
            formatArg = new StringEval("dd:mm:yyyy hh:mm:ss");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            testResult = new StringEval("16:11:1900 07:42:14");
            //Assert.AreEqual(testResult.ToString(), result.ToString());

            // this line is intended to compute how "November" would look like in the current locale
            String november = new SimpleDateFormat("MMMM").Format(new DateTime(2010, 11, 15), CultureInfo.CurrentCulture);
            
            // Again with Java style
            formatArg = new StringEval("MMMM dd, yyyy");
            args[1] = formatArg;
            //fix error in non-en Culture
            NPOI.SS.Formula.Functions.Text.Formatter = new NPOI.SS.UserModel.DataFormatter(CultureInfo.CurrentCulture);
            result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            testResult = new StringEval(november + " 16, 1900");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            // And Excel style
            formatArg = new StringEval("mmmm dd, yyyy");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, (short)-1);
            testResult = new StringEval(november + " 16, 1900");
            Assert.AreEqual(testResult.ToString(), result.ToString());
        }
    }

}