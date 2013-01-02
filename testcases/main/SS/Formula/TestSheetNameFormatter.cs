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

namespace TestCases.SS.Formula
{

    using NPOI.SS.Formula;
    using NUnit.Framework;
    using System;

    /**
     * Tests for {@link SheetNameFormatter}
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestSheetNameFormatter
    {

        private static void ConfirmFormat(String rawSheetName, String expectedSheetNameEncoding)
        {
            Assert.AreEqual(expectedSheetNameEncoding, SheetNameFormatter.Format(rawSheetName));
        }

        /**
         * Tests main public method 'format' 
         */
        [Test]
        public void TestFormat()
        {

            ConfirmFormat("abc", "abc");
            ConfirmFormat("123", "'123'");

            ConfirmFormat("my sheet", "'my sheet'"); // space
            ConfirmFormat("A:MEM", "'A:MEM'"); // colon

            ConfirmFormat("O'Brian", "'O''Brian'"); // single quote Gets doubled


            ConfirmFormat("3rdTimeLucky", "'3rdTimeLucky'"); // digit in first pos
            ConfirmFormat("_", "_"); // plain underscore OK
            ConfirmFormat("my_3rd_sheet", "my_3rd_sheet"); // underscores and digits OK
            ConfirmFormat("A12220", "'A12220'");
            ConfirmFormat("TAXRETURN19980415", "TAXRETURN19980415");
        }
        [Test]
        public void TestBooleanLiterals()
        {
            ConfirmFormat("TRUE", "'TRUE'");
            ConfirmFormat("FALSE", "'FALSE'");
            ConfirmFormat("True", "'True'");
            ConfirmFormat("fAlse", "'fAlse'");

            ConfirmFormat("Yes", "Yes");
            ConfirmFormat("No", "No");
        }

        private static void ConfirmCellNameMatch(String rawSheetName, bool expected)
        {
            Assert.AreEqual(expected, SheetNameFormatter.NameLooksLikePlainCellReference(rawSheetName));
        }

        /**
         * Tests functionality to determine whether a sheet name Containing only letters and digits
         * would look (to Excel) like a cell name.
         */
        [Test]
        public void TestLooksLikePlainCellReference()
        {

            ConfirmCellNameMatch("A1", true);
            ConfirmCellNameMatch("a111", true);
            ConfirmCellNameMatch("AA", false);
            ConfirmCellNameMatch("aa1", true);
            ConfirmCellNameMatch("A1A", false);
            ConfirmCellNameMatch("A1A1", false);
            ConfirmCellNameMatch("Sh3", false);
            ConfirmCellNameMatch("SALES20080101", false); // out of range
        }

        private static void ConfirmCellRange(String text, int numberOfPrefixLetters, bool expected)
        {
            String prefix = text.Substring(0, numberOfPrefixLetters);
            String suffix = text.Substring(numberOfPrefixLetters);
            Assert.AreEqual(expected, SheetNameFormatter.CellReferenceIsWithinRange(prefix, suffix));
        }

        /**
         * Tests exact boundaries for names that look very close to cell names (i.e. contain 1 or more
         * letters followed by one or more digits).
         */
        [Test]
        public void TestCellRange()
        {
            ConfirmCellRange("A1", 1, true);
            ConfirmCellRange("a111", 1, true);
            ConfirmCellRange("A65536", 1, true);
            ConfirmCellRange("A65537", 1, false);
            ConfirmCellRange("iv1", 2, true);
            ConfirmCellRange("IW1", 2, false);
            ConfirmCellRange("AAA1", 3, false);
            ConfirmCellRange("a111", 1, true);
            ConfirmCellRange("Sheet1", 6, false);
            ConfirmCellRange("iV65536", 2, true);  // max cell in Excel 97-2003
            ConfirmCellRange("IW65537", 2, false);
        }
    }

}