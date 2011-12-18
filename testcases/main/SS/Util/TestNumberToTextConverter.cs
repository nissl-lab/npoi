/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.SS.Util
{

    using System;
    using System.Text;
    using System.Collections.Generic;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.Util;
    /**
 * Tests for {@link NumberToTextConverter}
 * 
 * @author Josh Micich
 */
    [TestClass]
    public class TestNumberToTextConverter
    {
        /**
         * Confirms that <tt>ExcelNumberToTextConverter.toText(d)</tt> produces the right results.
         * As part of preparing this test class, the <tt>ExampleConversion</tt> instances should be set
         * up to contain the rendering as produced by Excel.
         */
        [TestMethod]
        public void TestAllNumberToText()
        {
            int failureCount = 0;

            ExampleConversion[] examples = NumberToTextConversionExamples.GetExampleConversions();

            for (int i = 0; i < examples.Length; i++)
            {
                ExampleConversion example = examples[i];
                try
                {
                    if (example.IsNaN())
                    {
                        ConfirmNaN(example.GetRawDoubleBits(), example.GetExcelRendering());
                        continue;
                    }
                    String actual = NumberToTextConverter.ToText(example.GetDoubleValue());
                    if (!example.GetExcelRendering().Equals(actual))
                    {
                        failureCount++;
                        String msg = "Error rendering for examples[" + i + "] "
                                + FormatExample(example) + " "
                                + " bad-result='" + actual + "' "
                                + "Excel String=" + example.GetExcelRendering();
                        //System.err.println(msg);
                        System.Console.WriteLine(msg);
                        continue;
                    }
                }
                catch (Exception e)
                {
                    failureCount++;
                    System.Console.WriteLine("Error in excel rendering for examples[" + i + "] "
                            + FormatExample(example) + "':" + e.Message);
                    //e.printStackTrace();
                    System.Console.Write(e.StackTrace);
                }
            }
            if (failureCount > 0)
            {
                throw new Exception(failureCount
                        + " error(s) in excel number to text conversion (see std-err)");
            }
        }

        private static String FormatExample(ExampleConversion example)
        {
            String hexLong = example.GetRawDoubleBits().ToString("X");
            String longRep = "0x" + "0000000000000000".Substring(hexLong.Length) + hexLong + "L";
            return "ec(" + longRep + ", \"" + example.GetJavaRendering() + "\", \"" + example.GetExcelRendering() + "\")";
        }

        /**
         * Excel's abnormal rendering of NaNs is both difficult to test and even reproduce in java. In
         * general, Excel does not attempt to use raw NaN in the IEEE sense. In {@link FormulaRecord}s,
         * Excel uses the NaN bit pattern to flag non-numeric (text, boolean, error) cached results.
         * If the formula result actually evaluates to raw NaN, Excel transforms it to <i>#NUM!</i>.
         * In other places (e.g. {@link NumberRecord}, {@link NumberPtg}, array items (via {@link 
         * ConstantValueParser}), there seems to be no special NaN translation scheme.  If a NaN bit 
         * pattern is somehow encoded into any of these places Excel actually attempts to render the 
         * values as a plain number. That is the unusual functionality that this method is testing.<p/>   
         * 
         * There are multiple encodings (bit patterns) for NaN, and CPUs and applications can convert
         * to a preferred NaN encoding  (Java prefers <tt>0x7FF8000000000000L</tt>).  Besides the 
         * special encoding in {@link FormulaRecord.SpecialCachedValue}, it is not known how/whether 
         * Excel attempts to encode NaN values.
         * 
         * Observed NaN behaviour on HotSpot/Windows:
         * <tt>Double.longBitsToDouble()</tt> will set one bit 51 (the NaN signaling flag) if it isn't
         *  already. <tt>Double.doubleToLongBits()</tt> will return a double with bit pattern 
         *  <tt>0x7FF8000000000000L</tt> for any NaN bit pattern supplied.<br/>
         * Differences are likely to be observed with other architectures.<p/>
         *  
         * <p/>
         * The few test case examples calling this method represent functionality which may not be 
         * important for POI to support.
         */
        private void ConfirmNaN(long l, String excelRep)
        {
            double d = BitConverter.Int64BitsToDouble(l);
            Assert.AreEqual("NaN", d.ToString());

            String strExcel = NumberToTextConverter.RawDoubleBitsToText(l);

            Assert.AreEqual(excelRep, strExcel);
        }
        [TestMethod]
        public void TestSimpleRendering_bug56156()
        {
            double dResult = 0.05 + 0.01; // values chosen to produce rounding anomaly
            String actualText = NumberToTextConverter.ToText(dResult);
            String jdkText = dResult.ToString();
            if (jdkText.Equals(actualText))
            {
                // "0.060000000000000005"
                throw new Exception("Should not use default JDK IEEE double rendering");
            }
            Assert.AreEqual("0.06", actualText);
        }
    }
}