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
    using System.Text;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;

    [TestFixture]
    public class TestProper
    {
        private ICell cell11;
        private IFormulaEvaluator Evaluator;

        [Test]
        public void TestValidHSSF()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            Evaluator = new HSSFFormulaEvaluator(wb);

            Confirm(wb);
        }

        [Test]
        public void TestValidXSSF()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            Evaluator = new XSSFFormulaEvaluator(wb);

            Confirm(wb);
        }

        private void Confirm(IWorkbook wb)
        {
            ISheet sheet = wb.CreateSheet("new sheet");
            cell11 = sheet.CreateRow(0).CreateCell(0);
            cell11.SetCellType(CellType.Formula);

            Confirm("PROPER(\"hi there\")", "Hi There"); //simple case
            Confirm("PROPER(\"what's up\")", "What'S Up"); //apostrophes are treated as word breaks
            Confirm("PROPER(\"I DON'T TH!NK SO!\")", "I Don'T Th!Nk So!"); //capitalization is ignored, special punctuation is treated as a word break
            Confirm("PROPER(\"dr\u00dcb\u00f6'\u00e4 \u00e9lo\u015f|\u00eb\u00e8 \")", "Dr\u00fcb\u00f6'\u00c4 \u00c9lo\u015f|\u00cb\u00e8 ");
            Confirm("PROPER(\"hi123 the123re\")", "Hi123 The123Re"); //numbers are treated as word breaks
            Confirm("PROPER(\"-\")", "-"); //nothing happens with ascii punctuation that is not upper or lower case
            Confirm("PROPER(\"!\u00a7$\")", "!\u00a7$"); //nothing happens with unicode punctuation (section sign) that is not upper or lower case
            Confirm("PROPER(\"/&%\")", "/&%"); //nothing happens with ascii punctuation that is not upper or lower case
            Confirm("PROPER(\"Apache POI\")", "Apache Poi"); //acronyms are not special
            Confirm("PROPER(\"  hello world\")", "  Hello World"); //leading whitespace is ignored

            String scharfes = "\u00df"; //German lowercase eszett, scharfes s, sharp s
            Confirm("PROPER(\"stra" + scharfes + "e\")", "Stra" + scharfes + "e");

            // CURRENTLY FAILS: result: "SSUnd"+scharfes
            // LibreOffice 5.0.3.2 behavior: "Sund"+scharfes
            // Excel 2013 behavior: ???
            //Confirm("PROPER(\"" + scharfes + "und" + scharfes + "\")", "SSund" + scharfes);
            Confirm("PROPER(\"" + scharfes + "und" + scharfes + "\")", "ßund" + scharfes); //should be ß in .net framework

            // also test longer string
            StringBuilder builder = new StringBuilder("A");
            StringBuilder expected = new StringBuilder("A");
            for (int i = 1; i < 254; i++)
            {
                builder.Append((char)(65 + (i % 26)));
                expected.Append((char)(97 + (i % 26)));
            }
            Confirm("PROPER(\"" + builder.ToString() + "\")", expected.ToString());
        }

        private void Confirm(String formulaText, String expectedResult)
        {
            cell11.CellFormula = (/*setter*/formulaText);
            Evaluator.ClearAllCachedResultValues();
            CellValue cv = Evaluator.Evaluate(cell11);
            if (cv.CellType != CellType.String)
            {
                Assert.Fail("Wrong result type: " + cv.FormatAsString());
            }
            String actualValue = cv.StringValue;
            Assert.AreEqual(expectedResult, actualValue);
        }

        [Test]
        public void Test()
        {
            checkProper("", "");
            checkProper("a", "A");
            checkProper("abc", "Abc");
            checkProper("abc abc", "Abc Abc");
            checkProper("abc/abc", "Abc/Abc");
            checkProper("ABC/ABC", "Abc/Abc");
            checkProper("aBc/ABC", "Abc/Abc");
            checkProper("aBc@#$%^&*()_+=-ABC", "Abc@#$%^&*()_+=-Abc");
            checkProper("aBc25aerg/ABC", "Abc25Aerg/Abc");
            checkProper("aBc/\u00C4\u00F6\u00DF\u00FC/ABC", "Abc/\u00C4\u00F6\u00DF\u00FC/Abc");  // Some German umlauts with uppercase first letter is not changed
            checkProper("\u00FC", "\u00DC");
            checkProper("\u00DC", "\u00DC");
            //checkProper("\u00DF", "SS");    // German "scharfes s" is uppercased to "SS"
            checkProper("\u00DF", "ß");  //should be ß in .net framework
            //checkProper("\u00DFomesing", "SSomesing");    // German "scharfes s" is uppercased to "SS"
            checkProper("\u00DFomesing", "ßomesing");   //should be ß in .net framework
            checkProper("aBc/\u00FC\u00C4\u00F6\u00DF\u00FC/ABC", "Abc/\u00DC\u00E4\u00F6\u00DF\u00FC/Abc");  // Some German umlauts with lowercase first letter is changed to uppercase
        }
        [Test]
        public void TestMicroBenchmark()
        {
            ValueEval strArg = new StringEval("some longer text that needs a number of replacements to check for runtime of different implementations");
            long start = TimeUtil.CurrentMillis();
            for (int i = 0; i < 300000; i++)
            {
                ValueEval ret = TextFunction.PROPER.Evaluate(new ValueEval[] { strArg }, 0, 0);
                Assert.AreEqual("Some Longer Text That Needs A Number Of Replacements To Check For Runtime Of Different Implementations", ((StringEval)ret).StringValue);
            }
            // Took aprox. 600ms on a decent Laptop in July 2016
            Console.WriteLine("Took: " + (TimeUtil.CurrentMillis() - start) + "ms");
        }
        private void checkProper(String input, String expected)
        {
            ValueEval strArg = new StringEval(input);
            ValueEval ret = TextFunction.PROPER.Evaluate(new ValueEval[] { strArg }, 0, 0);
            Assert.AreEqual(expected, ((StringEval)ret).StringValue);
        }

    }

}