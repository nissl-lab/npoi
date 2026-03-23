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

namespace TestCases.SS.Util
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS;
    using NPOI.SS.Util;
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;

    /// <summary>
    /// Tests <see cref="AreaReference.IsStructuredReference"/> and
    /// <see cref="AreaReference.CreateFromStructuredReference"/> for resolving
    /// Excel structured table references to concrete cell ranges.
    /// </summary>
    /// <remarks>
    /// Uses the <c>StructuredReferences.xlsx</c> sample workbook which contains
    /// a table named <c>\_Prime.1</c> on the "Table" sheet at A1:C7 with columns:
    /// <c>calc='#*'#</c>, <c>Name</c>, <c>Number</c>. No totals row. 6 data rows (rows 2-7).
    /// </remarks>
    [TestFixture]
    public class TestAreaReferenceStructuredReferences
    {
        [Test]
        public void IsStructuredReference_RecognizesTableReferences()
        {
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("Table1[#Headers]"));
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("Table1[#Data]"));
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("Table1[#All]"));
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("Table1[#Totals]"));
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("Table1[#This Row]"));
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("Table1[Column1]"));
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("Table1[[#Headers],[Column1]]"));
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("Table1[[#Data],[Col1]:[Col3]]"));
            ClassicAssert.IsTrue(AreaReference.IsStructuredReference("\\_Prime.1[#Headers]"));
        }

        [Test]
        public void IsStructuredReference_RejectsRegularReferences()
        {
            ClassicAssert.IsFalse(AreaReference.IsStructuredReference("A1"));
            ClassicAssert.IsFalse(AreaReference.IsStructuredReference("A1:B5"));
            ClassicAssert.IsFalse(AreaReference.IsStructuredReference("Sheet1!A1:B5"));
            ClassicAssert.IsFalse(AreaReference.IsStructuredReference("$A$1:$C$7"));
        }

        [Test]
        public void CreateFromStructuredReference_HeadersSpecifier()
        {
            // \_Prime.1 table is at A1:C7 on the "Table" sheet.
            // [#Headers] should resolve to the header row: row 0, columns 0-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            AreaReference area = AreaReference.CreateFromStructuredReference(
                "\\_Prime.1[#Headers]",
                SpreadsheetVersion.EXCEL2007,
                fpb);

            ClassicAssert.AreEqual("Table", area.FirstCell.SheetName);
            ClassicAssert.AreEqual(0, area.FirstCell.Row, "First row should be header row (0)");
            ClassicAssert.AreEqual(0, area.FirstCell.Col, "First column should be 0");
            ClassicAssert.AreEqual(0, area.LastCell.Row, "Last row should be header row (0)");
            ClassicAssert.AreEqual(2, area.LastCell.Col, "Last column should be 2");
        }

        [Test]
        public void CreateFromStructuredReference_DataSpecifier()
        {
            // [#Data] should resolve to data rows: rows 1-6, columns 0-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            AreaReference area = AreaReference.CreateFromStructuredReference(
                "\\_Prime.1[#Data]",
                SpreadsheetVersion.EXCEL2007,
                fpb);

            ClassicAssert.AreEqual(1, area.FirstCell.Row, "First data row should be 1");
            ClassicAssert.AreEqual(0, area.FirstCell.Col);
            ClassicAssert.AreEqual(6, area.LastCell.Row, "Last data row should be 6");
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void CreateFromStructuredReference_AllSpecifier()
        {
            // [#All] should resolve to entire table: rows 0-6, columns 0-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            AreaReference area = AreaReference.CreateFromStructuredReference(
                "\\_Prime.1[#All]",
                SpreadsheetVersion.EXCEL2007,
                fpb);

            ClassicAssert.AreEqual(0, area.FirstCell.Row, "Should start at row 0");
            ClassicAssert.AreEqual(0, area.FirstCell.Col);
            ClassicAssert.AreEqual(6, area.LastCell.Row, "Should end at row 6");
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void CreateFromStructuredReference_SingleColumnDataSpecifier()
        {
            // \_Prime.1[Number] should resolve to data rows of the "Number" column (index 2): rows 1-6, col 2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            AreaReference area = AreaReference.CreateFromStructuredReference(
                "\\_Prime.1[Number]",
                SpreadsheetVersion.EXCEL2007,
                fpb);

            ClassicAssert.AreEqual(1, area.FirstCell.Row);
            ClassicAssert.AreEqual(2, area.FirstCell.Col, "Number is column index 2");
            ClassicAssert.AreEqual(6, area.LastCell.Row);
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void CreateFromStructuredReference_HeadersWithColumn()
        {
            // \_Prime.1[[#Headers],[Number]] should resolve to header cell of "Number" column: row 0, col 2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            AreaReference area = AreaReference.CreateFromStructuredReference(
                "\\_Prime.1[[#Headers],[Number]]",
                SpreadsheetVersion.EXCEL2007,
                fpb);

            ClassicAssert.AreEqual(0, area.FirstCell.Row);
            ClassicAssert.AreEqual(2, area.FirstCell.Col);
            ClassicAssert.AreEqual(0, area.LastCell.Row);
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void CreateFromStructuredReference_ThisRowSpecifier()
        {
            // [#This Row] at row index 3 should resolve to row 3, columns 0-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            AreaReference area = AreaReference.CreateFromStructuredReference(
                "\\_Prime.1[#This Row]",
                SpreadsheetVersion.EXCEL2007,
                fpb,
                rowIndex: 3);

            ClassicAssert.AreEqual(3, area.FirstCell.Row, "Should resolve to the specified row index");
            ClassicAssert.AreEqual(0, area.FirstCell.Col);
            ClassicAssert.AreEqual(3, area.LastCell.Row);
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void CreateFromStructuredReference_ColumnRange()
        {
            // \_Prime.1[[Name]:[Number]] should resolve to data rows of columns 1-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            AreaReference area = AreaReference.CreateFromStructuredReference(
                "\\_Prime.1[[Name]:[Number]]",
                SpreadsheetVersion.EXCEL2007,
                fpb);

            ClassicAssert.AreEqual(1, area.FirstCell.Row, "Data starts at row 1");
            ClassicAssert.AreEqual(1, area.FirstCell.Col, "Name is column 1");
            ClassicAssert.AreEqual(6, area.LastCell.Row, "Data ends at row 6");
            ClassicAssert.AreEqual(2, area.LastCell.Col, "Number is column 2");
        }

        [Test]
        public void CreateFromStructuredReference_NullWorkbook_Throws()
        {
            Assert.Throws<ArgumentNullException>(() =>
                AreaReference.CreateFromStructuredReference(
                    "Table1[#Headers]",
                    SpreadsheetVersion.EXCEL2007,
                    null));
        }

        [Test]
        public void CreateFromStructuredReference_InvalidTableName_Throws()
        {
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            // BaseXSSFEvaluationWorkbook.GetTable() throws KeyNotFoundException for unknown tables
            Assert.Throws<KeyNotFoundException>(() =>
                AreaReference.CreateFromStructuredReference(
                    "NonExistentTable[#Headers]",
                    SpreadsheetVersion.EXCEL2007,
                    fpb));
        }

        [Test]
        public void AreaReferenceConstructor_StillThrowsOnStructuredRef()
        {
            // Verify that the existing constructor still throws on structured references,
            // so callers know they need to use CreateFromStructuredReference instead.
            Assert.Throws<ArgumentException>(() =>
                new AreaReference("Table1[#Headers]", SpreadsheetVersion.EXCEL2007));
        }
    }
}
