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
namespace TestCases.XSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    [TestFixture]
    public class TestXSSFDataValidationConstraint
    {
        static int listType = ValidationType.LIST;
        static int ignoredType = OperatorType.IGNORED;

        // See bug 59719
        [Test]
        public void ListLiteralsQuotesAreStripped_formulaConstructor()
        {
            // literal list, using formula constructor
            String literal = "\"one, two, three\"";
            String[] expected = new String[] { "one", "two", "three" };
            IDataValidationConstraint constraint = new XSSFDataValidationConstraint(listType, ignoredType, literal, null);

            CollectionAssert.AreEqual(expected, constraint.ExplicitListValues);
            // Excel and DataValidationConstraint Parser ignore (strip) whitespace; quotes should still be intact
            // FIXME: whitespace wasn't stripped
            ClassicAssert.AreEqual(literal, constraint.Formula1);
        }

        [Test]
        public void ListLiteralsQuotesAreStripped_arrayConstructor()
        {
            // literal list, using array constructor
            String literal = "\"one, two, three\"";
            String[] expected = new String[] { "one", "two", "three" };
            IDataValidationConstraint constraint = new XSSFDataValidationConstraint(expected);
            CollectionAssert.AreEqual(expected, constraint.ExplicitListValues);
            // Excel and DataValidationConstraint Parser ignore (strip) whitespace; quotes should still be intact
            ClassicAssert.AreEqual(literal.Replace(" ", ""), constraint.Formula1);
        }

        // See issue 1369
        [Test]
        public void ListLiteralsWithSpecialCharactersAreWrittenAsExcelStrings()
        {
            String[] values = new String[] { "a\"b", "x&y", "<b>", "it's", "a]]>b" };
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IDataValidationHelper helper = sheet.GetDataValidationHelper();
            sheet.AddValidationData(helper.CreateValidation(helper.CreateExplicitListConstraint(values), new CellRangeAddressList(0, 0, 0, 0)));

            MemoryStream output = new MemoryStream();
            workbook.Write(output, true);

            XSSFWorkbook readBack = new XSSFWorkbook(new MemoryStream(output.ToArray()));
            IDataValidationConstraint readConstraint = readBack.GetSheetAt(0).GetDataValidations()[0].ValidationConstraint;
            ClassicAssert.AreEqual("\"a\"\"b,x&y,<b>,it's,a]]>b\"", readConstraint.Formula1);
            CollectionAssert.AreEqual(values, readConstraint.ExplicitListValues);
        }

        [Test]
        public void RangeReference()
        {
            // (unnamed range) reference list        
            String reference = "A1:A5";
            IDataValidationConstraint constraint = new XSSFDataValidationConstraint(listType, ignoredType, reference, null);
            ClassicAssert.IsNull(constraint.ExplicitListValues);
            ClassicAssert.AreEqual("A1:A5", constraint.Formula1);
        }

        [Test]
        public void NamedRangeReference()
        {
            // named range list
            String namedRange = "MyNamedRange";
            IDataValidationConstraint constraint = new XSSFDataValidationConstraint(listType, ignoredType, namedRange, null);
            ClassicAssert.IsNull(constraint.ExplicitListValues);
            ClassicAssert.AreEqual("MyNamedRange", constraint.Formula1);
        }

    }

}