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
    /// Tests the <see cref="AreaReference"/> constructor overload that accepts an
    /// <see cref="NPOI.SS.Formula.IFormulaParsingWorkbook"/> for resolving Excel
    /// structured table references to concrete cell ranges.
    /// </summary>
    /// <remarks>
    /// Uses the <c>StructuredReferences.xlsx</c> sample workbook which contains
    /// a table named <c>\_Prime.1</c> on the "Table" sheet at A1:C7 with columns:
    /// <c>calc='#*'#</c>, <c>Name</c>, <c>Number</c>. No totals row. 6 data rows (rows 2-7).
    /// </remarks>
    [TestFixture]
    public class TestAreaReferenceStructuredReferences
    {
        // ---- IsStructuredReference detection ----

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

        // ---- Constructor with workbook: structured reference specifiers ----

        [Test]
        public void Constructor_WithWorkbook_HeadersSpecifier()
        {
            // \_Prime.1 table is at A1:C7 on the "Table" sheet.
            // [#Headers] should resolve to the header row: row 0, columns 0-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("\\_Prime.1[#Headers]", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual("Table", area.FirstCell.SheetName);
            ClassicAssert.AreEqual(0, area.FirstCell.Row, "First row should be header row (0)");
            ClassicAssert.AreEqual(0, area.FirstCell.Col, "First column should be 0");
            ClassicAssert.AreEqual(0, area.LastCell.Row, "Last row should be header row (0)");
            ClassicAssert.AreEqual(2, area.LastCell.Col, "Last column should be 2");
        }

        [Test]
        public void Constructor_WithWorkbook_DataSpecifier()
        {
            // [#Data] should resolve to data rows: rows 1-6, columns 0-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("\\_Prime.1[#Data]", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual(1, area.FirstCell.Row, "First data row should be 1");
            ClassicAssert.AreEqual(0, area.FirstCell.Col);
            ClassicAssert.AreEqual(6, area.LastCell.Row, "Last data row should be 6");
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void Constructor_WithWorkbook_AllSpecifier()
        {
            // [#All] should resolve to entire table: rows 0-6, columns 0-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("\\_Prime.1[#All]", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual(0, area.FirstCell.Row, "Should start at row 0");
            ClassicAssert.AreEqual(0, area.FirstCell.Col);
            ClassicAssert.AreEqual(6, area.LastCell.Row, "Should end at row 6");
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void Constructor_WithWorkbook_SingleColumnDataSpecifier()
        {
            // \_Prime.1[Number] should resolve to data rows of the "Number" column (index 2): rows 1-6, col 2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("\\_Prime.1[Number]", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual(1, area.FirstCell.Row);
            ClassicAssert.AreEqual(2, area.FirstCell.Col, "Number is column index 2");
            ClassicAssert.AreEqual(6, area.LastCell.Row);
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void Constructor_WithWorkbook_HeadersWithColumn()
        {
            // \_Prime.1[[#Headers],[Number]] should resolve to header cell of "Number" column: row 0, col 2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("\\_Prime.1[[#Headers],[Number]]", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual(0, area.FirstCell.Row);
            ClassicAssert.AreEqual(2, area.FirstCell.Col);
            ClassicAssert.AreEqual(0, area.LastCell.Row);
            ClassicAssert.AreEqual(2, area.LastCell.Col);
            ClassicAssert.IsTrue(area.IsSingleCell, "Header of a single column should be a single cell");
        }

        [Test]
        public void Constructor_WithWorkbook_ThisRowSpecifier()
        {
            // [#This Row] at row index 3 should resolve to row 3, columns 0-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("\\_Prime.1[#This Row]", SpreadsheetVersion.EXCEL2007, fpb, 3);

            ClassicAssert.AreEqual(3, area.FirstCell.Row, "Should resolve to the specified row index");
            ClassicAssert.AreEqual(0, area.FirstCell.Col);
            ClassicAssert.AreEqual(3, area.LastCell.Row);
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void Constructor_WithWorkbook_ColumnRange()
        {
            // \_Prime.1[[Name]:[Number]] should resolve to data rows of columns 1-2.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("\\_Prime.1[[Name]:[Number]]", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual(1, area.FirstCell.Row, "Data starts at row 1");
            ClassicAssert.AreEqual(1, area.FirstCell.Col, "Name is column 1");
            ClassicAssert.AreEqual(6, area.LastCell.Row, "Data ends at row 6");
            ClassicAssert.AreEqual(2, area.LastCell.Col, "Number is column 2");
        }

        [Test]
        public void Constructor_WithWorkbook_InvalidTableName_Throws()
        {
            // The constructor can throw several exception types for invalid structured references:
            // - KeyNotFoundException: table name not found in workbook (from GetTable())
            // - FormulaParseException: malformed structured reference syntax
            // - InvalidOperationException: reference resolves to something other than a single area
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            Assert.Throws<KeyNotFoundException>(() =>
                new AreaReference("NonExistentTable[#Headers]", SpreadsheetVersion.EXCEL2007, fpb));
        }

        // ---- Constructor without workbook: backward compatibility ----

        [Test]
        public void Constructor_WithoutWorkbook_StructuredRefStillThrows()
        {
            // The 2-param constructor (no workbook) must still throw on structured references,
            // preserving backward-compatible behavior.
            Assert.Throws<ArgumentException>(() =>
                new AreaReference("Table1[#Headers]", SpreadsheetVersion.EXCEL2007));
        }

        [Test]
        public void Constructor_WithoutWorkbook_RegularRefStillWorks()
        {
            // Verify the 2-param constructor still handles all standard reference types
            // after the refactor to chain through the 4-param constructor.
            var singleCell = new AreaReference("B3", SpreadsheetVersion.EXCEL2007);
            ClassicAssert.AreEqual(2, singleCell.FirstCell.Row);
            ClassicAssert.AreEqual(1, singleCell.FirstCell.Col);
            ClassicAssert.IsTrue(singleCell.IsSingleCell);

            var multiCell = new AreaReference("A1:C7", SpreadsheetVersion.EXCEL2007);
            ClassicAssert.AreEqual(0, multiCell.FirstCell.Row);
            ClassicAssert.AreEqual(0, multiCell.FirstCell.Col);
            ClassicAssert.AreEqual(6, multiCell.LastCell.Row);
            ClassicAssert.AreEqual(2, multiCell.LastCell.Col);
            ClassicAssert.IsFalse(multiCell.IsSingleCell);

            var withSheet = new AreaReference("Sheet1!A1:B2", SpreadsheetVersion.EXCEL2007);
            ClassicAssert.AreEqual("Sheet1", withSheet.FirstCell.SheetName);
            ClassicAssert.AreEqual(0, withSheet.FirstCell.Row);
            ClassicAssert.AreEqual(1, withSheet.LastCell.Col);
        }

        [Test]
        public void Constructor_WithoutWorkbook_WholeColumnRefStillWorks()
        {
            // Whole-column references (e.g., "A:C") go through a special parsing path.
            var wholeCol = new AreaReference("A:C", SpreadsheetVersion.EXCEL97);
            ClassicAssert.AreEqual(0, wholeCol.FirstCell.Col);
            ClassicAssert.AreEqual(2, wholeCol.LastCell.Col);
            ClassicAssert.IsTrue(wholeCol.IsWholeColumnReference());
        }

        [Test]
        public void Constructor_WithoutWorkbook_AbsoluteRefStillWorks()
        {
            var abs = new AreaReference("$A$1:$C$7", SpreadsheetVersion.EXCEL2007);
            ClassicAssert.AreEqual(0, abs.FirstCell.Row);
            ClassicAssert.AreEqual(0, abs.FirstCell.Col);
            ClassicAssert.IsTrue(abs.FirstCell.IsRowAbsolute);
            ClassicAssert.IsTrue(abs.FirstCell.IsColAbsolute);
            ClassicAssert.AreEqual(6, abs.LastCell.Row);
            ClassicAssert.AreEqual(2, abs.LastCell.Col);
        }

        // ---- Constructor with workbook: regular references pass through ----

        [Test]
        public void Constructor_WithWorkbook_RegularRefPassesThrough()
        {
            // When a workbook is provided but the reference is a normal A1-style ref,
            // it should work exactly like the 2-param constructor.
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("A1:C7", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual(0, area.FirstCell.Row);
            ClassicAssert.AreEqual(0, area.FirstCell.Col);
            ClassicAssert.AreEqual(6, area.LastCell.Row);
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }

        [Test]
        public void Constructor_WithWorkbook_SheetQualifiedRefPassesThrough()
        {
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("Table!A1:C7", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual("Table", area.FirstCell.SheetName);
            ClassicAssert.AreEqual(0, area.FirstCell.Row);
            ClassicAssert.AreEqual(6, area.LastCell.Row);
        }

        [Test]
        public void Constructor_WithWorkbook_SingleCellRefPassesThrough()
        {
            using XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("StructuredReferences.xlsx");
            var fpb = XSSFEvaluationWorkbook.Create(wb);

            var area = new AreaReference("B3", SpreadsheetVersion.EXCEL2007, fpb);

            ClassicAssert.AreEqual(2, area.FirstCell.Row);
            ClassicAssert.AreEqual(1, area.FirstCell.Col);
            ClassicAssert.IsTrue(area.IsSingleCell);
        }

        [Test]
        public void Constructor_NullWorkbook_RegularRefStillWorks()
        {
            // Explicitly passing null workbook should behave identically to 2-param constructor.
            var area = new AreaReference("A1:C7", SpreadsheetVersion.EXCEL2007, null);

            ClassicAssert.AreEqual(0, area.FirstCell.Row);
            ClassicAssert.AreEqual(0, area.FirstCell.Col);
            ClassicAssert.AreEqual(6, area.LastCell.Row);
            ClassicAssert.AreEqual(2, area.LastCell.Col);
        }
    }
}
