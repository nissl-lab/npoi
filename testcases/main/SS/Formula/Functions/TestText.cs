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
    using System.Collections.Generic;
    using NPOI.SS.UserModel;

    /**
     * Test case for TEXT()
     *
     * @author Stephen Wolke (smwolke at geistig.com)
     */
    [TestFixture]
    public class TestText
    {
        private readonly List<ErrorEval> EXCEL_ERRORS = new(11) {
            ErrorEval.NULL_INTERSECTION,
            ErrorEval.DIV_ZERO,
            ErrorEval.VALUE_INVALID,
            ErrorEval.REF_INVALID,
            ErrorEval.NAME_INVALID,
            ErrorEval.NUM_ERROR,
            ErrorEval.NA
        };

        private readonly CultureInfo _currentCulture = CultureInfo.InvariantCulture;

        [OneTimeSetUp]
        public void SetUp()
        {
            Text.Formatter = new DataFormatter(_currentCulture);
        }

        //private static TextFunction T = null;
        [Test]
        public void TestTextWithStringFirstArg()
        {
            ValueEval strArg = new StringEval("abc");
            ValueEval formatArg = new StringEval("abc");
            ValueEval[] args = { strArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, -1);
            Assert.AreEqual(strArg.ToString(), result.ToString());
        }

        [Test]
        public void TestTextWithDecimalFormatSecondArg()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            ValueEval numArg = new NumberEval(321321.321);
            ValueEval formatArg = new StringEval("#,###.00000");
            ValueEval[] args = { numArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, -1);

            NumberFormatInfo fs = _currentCulture.NumberFormat;
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
            result = TextFunction.TEXT.Evaluate(args, -1, -1);
            testResult = new StringEval("$321" + decimalSeparator + "3");
            Assert.AreEqual(testResult.ToString(), result.ToString());
        }

        [Test]
        public void TestTextWithFractionFormatSecondArg()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            ValueEval numArg = new NumberEval(321.321);
            ValueEval formatArg = new StringEval("# #/#");
            ValueEval[] args = { numArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, -1);
            ValueEval testResult = new StringEval("321 1/3");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            formatArg = new StringEval("# #/##");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, -1);
            testResult = new StringEval("321 26/81");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            formatArg = new StringEval("#/##");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, -1);
            testResult = new StringEval("26027/81");
            Assert.AreEqual(testResult.ToString(), result.ToString());
        }

        [Test]
        public void TestTextWithDateFormatSecondArg()
        {
            ValueEval numArg = new NumberEval(321.321);
            ValueEval formatArg = new StringEval("dd:MM:yyyy hh:mm:ss");
            ValueEval[] args = { numArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, -1);
            ValueEval testResult = new StringEval("16:11:1900 07:42:14");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            // Excel also supports "m before h is month"
            formatArg = new StringEval("dd:mm:yyyy hh:mm:ss");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, -1);
            testResult = new StringEval("16:11:1900 07:42:14");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            // this line is intended to compute how "November" would look like in the current locale
            string november = new SimpleDateFormat("MMMM").Format(new DateTime(2010, 11, 15), CultureInfo.CurrentCulture);
            
            // Again with Java style
            formatArg = new StringEval("MMMM dd, yyyy");
            args[1] = formatArg;

            result = TextFunction.TEXT.Evaluate(args, -1, -1);
            testResult = new StringEval(november + " 16, 1900");
            Assert.AreEqual(testResult.ToString(), result.ToString());

            // And Excel style
            formatArg = new StringEval("mmmm dd, yyyy");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, -1);
            testResult = new StringEval(november + " 16, 1900");
            Assert.AreEqual(testResult.ToString(), result.ToString());
        }

        [Test]
        public void TestTextWithISODateTimeFormatSecondArg()
        {
            ValueEval numArg = new NumberEval(321.321);
            ValueEval formatArg = new StringEval("yyyy-mm-ddThh:MM:ss");
            ValueEval[] args = { numArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, -1);
            ValueEval testResult = new StringEval("1900-11-16T07:42:14");
            Assert.AreEqual(testResult.ToString(), result.ToString());
            
 	        // test milliseconds
             formatArg = new StringEval("yyyy-mm-ddThh:MM:ss.000");
            args[1] = formatArg;
            result = TextFunction.TEXT.Evaluate(args, -1, -1);
            testResult = new StringEval("1900-11-16T07:42:14.400");
            Assert.AreEqual(testResult.ToString(), result.ToString());
        }

        // Test cases from the workbook attached to the bug 67475 which were OK

        [Test]
        public void TestTextVariousValidNumberFormats()
        {
            // negative values: 3 decimals
            Testtext(new NumberEval(-123456.789012345), new StringEval("#0.000"), "-123456.789");
            // no decimals
            Testtext(new NumberEval(-123456.789012345), new StringEval("000000"), "-123457");
            // common format - more digits
            Testtext(new NumberEval(-123456.789012345), new StringEval("00.0000"), "-123456.7890");
            // common format - less digits
            Testtext(new NumberEval(-12.78), new StringEval("00000.000000"), "-00012.780000");
            // half up
            Testtext(new NumberEval(-0.56789012375), new StringEval("#0.0000000000"), "-0.5678901238");
            // half up
            Testtext(new NumberEval(-0.56789012385), new StringEval("#0.0000000000"), "-0.5678901239");
            // positive values: 3 decimals
            Testtext(new NumberEval(123456.789012345), new StringEval("#0.000"), "123456.789");
            // no decimals
            Testtext(new NumberEval(123456.789012345), new StringEval("000000"), "123457");
            // common format - more digits
            Testtext(new NumberEval(123456.789012345), new StringEval("00.0000"), "123456.7890");
            // common format - less digits
            Testtext(new NumberEval(12.78), new StringEval("00000.000000"), "00012.780000");
            // half up
            Testtext(new NumberEval(0.56789012375), new StringEval("#0.0000000000"), "0.5678901238");
            // half up
            Testtext(new NumberEval(0.56789012385), new StringEval("#0.0000000000"), "0.5678901239");
        }

        [Test]
 	    public void testTextBlankTreatedAsZero()
        {
            Testtext(BlankEval.instance, new StringEval("#0.000"), "0.000");
        }
 	
 	    [Test]
 	    public void testTextStrangeFormat()
        {
            // number 0
            Testtext(new NumberEval(-123456.789012345), new NumberEval(0), "-123457");
            // negative number with few zeros
            Testtext(new NumberEval(-123456.789012345), new NumberEval(-0.0001), "--123456.7891");
            // format starts with "."
            Testtext(new NumberEval(0.0123), new StringEval(".000"), ".012");
            // one zero negative
            Testtext(new NumberEval(1001.202), new NumberEval(-8808), "-8810018");
            // format contains 0
            Testtext(new NumberEval(43368.0), new NumberEval(909), "9433689");
        }
 	
 	    [Test]
 	    public void TestTextErrorAsFormat()
        {
            foreach(ErrorEval errorEval in EXCEL_ERRORS)
            {
                Testtext(new NumberEval(3.14), errorEval, errorEval);
                Testtext(BoolEval.TRUE, errorEval, errorEval);
                Testtext(BoolEval.FALSE, errorEval, errorEval);
            }
        }
 	
 	    [Test]
 	    public void TestTextErrorAsValue()
        {
            foreach(ErrorEval errorEval in EXCEL_ERRORS)
            {
                Testtext(errorEval, new StringEval("#0.000"), errorEval);
                Testtext(errorEval, new StringEval("yyyymmmdd"), errorEval);
            }
        }
 	
 	    // Test cases from the workbook attached to the bug 67475 which were failing and are fixed by the patch
 	
 	    [Test]
        public void TestTextEmptyStringWithDateFormat()
        {
            Testtext(new StringEval(""), new StringEval("yyyymmmdd"), "");
        }
 	
 	    [Test]
        public void TestTextAnyTextWithDateFormat()
        {
            Testtext(new StringEval("anyText"), new StringEval("yyyymmmdd"), "anyText");
        }
 	
 	    [Test]
        public void TestTextBooleanWithDateFormat()
        {
            Testtext(BoolEval.TRUE, new StringEval("yyyymmmdd"), BoolEval.TRUE.StringValue);
            Testtext(BoolEval.FALSE, new StringEval("yyyymmmdd"), BoolEval.FALSE.StringValue);
        }
 	
 	    [Test]
 	    public void TestTextNumberWithBooleanFormat()
        {
            Testtext(new NumberEval(43368), BoolEval.TRUE, ErrorEval.VALUE_INVALID);
            Testtext(new NumberEval(43368), BoolEval.FALSE, ErrorEval.VALUE_INVALID);

            Testtext(new NumberEval(3.14), BoolEval.TRUE, ErrorEval.VALUE_INVALID);
            Testtext(new NumberEval(3.14), BoolEval.FALSE, ErrorEval.VALUE_INVALID);
        }

        [Test]
        public void TestTextEmptyStringWithNumberFormat()
        {
            Testtext(new StringEval(""), new StringEval("#0.000"), "");
        }
 	
 	    [Test]
 	    public void TestTextAnyTextWithNumberFormat()
        {
            Testtext(new StringEval("anyText"), new StringEval("#0.000"), "anyText");
        }

        [Test]
        public void TestTextBooleanWithNumberFormat()
        {
            Testtext(BoolEval.TRUE, new StringEval("#0.000"), BoolEval.TRUE.StringValue);
            Testtext(BoolEval.FALSE, new StringEval("#0.000"), BoolEval.FALSE.StringValue);
        }

        private static void Testtext(ValueEval valueArg, ValueEval formatArg, string expectedResult)
        {
            ValueEval[] args = { valueArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, -1);

            Assert.IsTrue(result is StringEval, "Expected StringEval got " + result.GetType().Name);
            Assert.AreEqual(expectedResult, ((StringEval) result).StringValue);
        }
 	
 	    private static void Testtext(ValueEval valueArg, ValueEval formatArg, ErrorEval expectedResult)
        {
            ValueEval[] args = { valueArg, formatArg };
            ValueEval result = TextFunction.TEXT.Evaluate(args, -1, -1);

            Assert.IsTrue(result is ErrorEval, "Expected ErrorEval got " + result.GetType().Name);
            Assert.AreEqual(expectedResult, result);
        }
    }
}