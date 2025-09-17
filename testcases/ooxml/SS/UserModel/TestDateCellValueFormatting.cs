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

using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;

namespace TestCases.SS.UserModel
{
    /// <summary>
    /// Tests for Number cell wrongly detected as DateTime
    /// Ensures DateCellValue only returns values for properly date-formatted cells
    /// </summary>
    [TestFixture]
    public class TestDateCellValueFormatting
    {
        /// <summary>
        /// Test that numeric cells without date formatting should return null for DateCellValue
        /// This addresses the core bug where numeric values like 12345.67 were being
        /// converted to DateTime (1933-10-18T16:04:48) instead of remaining as numbers
        /// </summary>
        [Test]
        public void TestNumericCellNotDateFormatted_ShouldNotReturnDateCellValue()
        {
            // Test with HSSF (XLS format)
            using (var hssfWorkbook = new HSSFWorkbook())
            {
                var hssfSheet = hssfWorkbook.CreateSheet("Test");
                var hssfRow = hssfSheet.CreateRow(0);
                var hssfCell = hssfRow.CreateCell(0);

                // Set a numeric value without date formatting
                hssfCell.SetCellValue(12345.67);

                // Verify it's a numeric cell
                Assert.AreEqual(CellType.Numeric, hssfCell.CellType);
                Assert.AreEqual(12345.67, hssfCell.NumericCellValue, 0.001);

                // Verify it's NOT date formatted
                Assert.IsFalse(DateUtil.IsCellDateFormatted(hssfCell));

                // The fix: DateCellValue should return null for non-date-formatted numeric cells
                Assert.IsNull(hssfCell.DateCellValue, "Numeric cell without date formatting should return null for DateCellValue");
            }

            // Test with XSSF (XLSX format)
            using (var xssfWorkbook = new XSSFWorkbook())
            {
                var xssfSheet = xssfWorkbook.CreateSheet("Test");
                var xssfRow = xssfSheet.CreateRow(0);
                var xssfCell = xssfRow.CreateCell(0);

                // Set a numeric value without date formatting
                xssfCell.SetCellValue(12345.67);

                // Verify it's a numeric cell
                Assert.AreEqual(CellType.Numeric, xssfCell.CellType);
                Assert.AreEqual(12345.67, xssfCell.NumericCellValue, 0.001);

                // Verify it's NOT date formatted
                Assert.IsFalse(DateUtil.IsCellDateFormatted(xssfCell));

                // The fix: DateCellValue should return null for non-date-formatted numeric cells
                Assert.IsNull(xssfCell.DateCellValue, "Numeric cell without date formatting should return null for DateCellValue");
            }
        }

        /// <summary>
        /// Test that formula cells with numeric results without date formatting should return null for DateCellValue
        /// This addresses the bug where formula results like 123.45 were being
        /// converted to DateTime (1900-05-02T10:48:00) instead of remaining as numbers
        /// </summary>
        [Test]
        public void TestFormulaCellNotDateFormatted_ShouldNotReturnDateCellValue()
        {
            // Test with HSSF (XLS format)
            using (var hssfWorkbook = new HSSFWorkbook())
            {
                var hssfSheet = hssfWorkbook.CreateSheet("Test");
                var hssfRow = hssfSheet.CreateRow(0);

                // Create input cells
                hssfRow.CreateCell(0).SetCellValue(100);
                hssfRow.CreateCell(1).SetCellValue(23.45);

                // Create formula cell
                var hssfFormulaCell = hssfRow.CreateCell(2);
                hssfFormulaCell.SetCellFormula("A1+B1");

                // Evaluate the formula
                var evaluator = hssfWorkbook.GetCreationHelper().CreateFormulaEvaluator();
                evaluator.EvaluateFormulaCell(hssfFormulaCell);

                // Verify it's a formula cell with numeric result
                Assert.AreEqual(CellType.Formula, hssfFormulaCell.CellType);
                Assert.AreEqual(CellType.Numeric, hssfFormulaCell.CachedFormulaResultType);
                Assert.AreEqual(123.45, hssfFormulaCell.NumericCellValue, 0.001);

                // Verify it's NOT date formatted
                Assert.IsFalse(DateUtil.IsCellDateFormatted(hssfFormulaCell));

                // The fix: DateCellValue should return null for non-date-formatted formula cells
                Assert.IsNull(hssfFormulaCell.DateCellValue, "Formula cell without date formatting should return null for DateCellValue");
            }

            // Test with XSSF (XLSX format)
            using (var xssfWorkbook = new XSSFWorkbook())
            {
                var xssfSheet = xssfWorkbook.CreateSheet("Test");
                var xssfRow = xssfSheet.CreateRow(0);

                // Create input cells
                xssfRow.CreateCell(0).SetCellValue(100);
                xssfRow.CreateCell(1).SetCellValue(23.45);

                // Create formula cell
                var xssfFormulaCell = xssfRow.CreateCell(2);
                xssfFormulaCell.SetCellFormula("A1+B1");

                // Evaluate the formula
                var evaluator = xssfWorkbook.GetCreationHelper().CreateFormulaEvaluator();
                evaluator.EvaluateFormulaCell(xssfFormulaCell);

                // Verify it's a formula cell with numeric result
                Assert.AreEqual(CellType.Formula, xssfFormulaCell.CellType);
                Assert.AreEqual(CellType.Numeric, xssfFormulaCell.CachedFormulaResultType);
                Assert.AreEqual(123.45, xssfFormulaCell.NumericCellValue, 0.001);

                // Verify it's NOT date formatted
                Assert.IsFalse(DateUtil.IsCellDateFormatted(xssfFormulaCell));

                // The fix: DateCellValue should return null for non-date-formatted formula cells
                Assert.IsNull(xssfFormulaCell.DateCellValue, "Formula cell without date formatting should return null for DateCellValue");
            }
        }

        /// <summary>
        /// Test that properly date-formatted cells should continue to work correctly
        /// This ensures our fix doesn't break existing date functionality
        /// </summary>
        [Test]
        public void TestActualDateCells_ShouldReturnDateCellValue()
        {
            var testDate = new DateTime(2023, 12, 25, 14, 30, 45);

            // Test with HSSF (XLS format)
            using (var hssfWorkbook = new HSSFWorkbook())
            {
                var hssfSheet = hssfWorkbook.CreateSheet("Test");
                var hssfRow = hssfSheet.CreateRow(0);
                var hssfCell = hssfRow.CreateCell(0);

                // Set a DateTime value (this automatically applies date formatting)
                hssfCell.SetCellValue(testDate);

                // Verify it's a numeric cell (dates are stored as numbers)
                Assert.AreEqual(CellType.Numeric, hssfCell.CellType);

                // Verify it IS date formatted
                Assert.IsTrue(DateUtil.IsCellDateFormatted(hssfCell));

                // DateCellValue should work correctly for properly formatted date cells
                Assert.IsNotNull(hssfCell.DateCellValue);
                Assert.AreEqual(testDate, hssfCell.DateCellValue);
            }

            // Test with XSSF (XLSX format)
            using (var xssfWorkbook = new XSSFWorkbook())
            {
                var xssfSheet = xssfWorkbook.CreateSheet("Test");
                var xssfRow = xssfSheet.CreateRow(0);
                var xssfCell = xssfRow.CreateCell(0);

                // Set a DateTime value (this automatically applies date formatting)
                xssfCell.SetCellValue(testDate);

                // Verify it's a numeric cell (dates are stored as numbers)
                Assert.AreEqual(CellType.Numeric, xssfCell.CellType);

                // Verify it IS date formatted
                Assert.IsTrue(DateUtil.IsCellDateFormatted(xssfCell));

                // DateCellValue should work correctly for properly formatted date cells
                Assert.IsNotNull(xssfCell.DateCellValue);
                Assert.AreEqual(testDate, xssfCell.DateCellValue);
            }
        }

        /// <summary>
        /// Test edge cases with various numeric values that could be misinterpreted as dates
        /// </summary>
        [Test]
        public void TestEdgeCases_VariousNumericValues()
        {
            var testValues = new double[] {
                0.0,        // Could be 1900-01-00 (invalid date)
                1.0,        // Could be 1900-01-01
                365.0,      // Could be 1900-12-31
                44197.0,    // Could be 2021-01-01 in Excel
                -1234.5,    // Negative number
                999999.99   // Large number
            };

            // Test with HSSF (XLS format)
            using (var hssfWorkbook = new HSSFWorkbook())
            {
                var hssfSheet = hssfWorkbook.CreateSheet("Test");

                for (int i = 0; i < testValues.Length; i++)
                {
                    var hssfRow = hssfSheet.CreateRow(i);
                    var hssfCell = hssfRow.CreateCell(0);

                    // Set numeric value without date formatting
                    hssfCell.SetCellValue(testValues[i]);

                    // Verify it's NOT date formatted
                    Assert.IsFalse(DateUtil.IsCellDateFormatted(hssfCell),
                        $"Value {testValues[i]} should not be date formatted");

                    // DateCellValue should return null
                    Assert.IsNull(hssfCell.DateCellValue,
                        $"Numeric value {testValues[i]} should return null for DateCellValue");

                    // But numeric access should still work
                    Assert.AreEqual(testValues[i], hssfCell.NumericCellValue, 0.001,
                        $"Numeric value {testValues[i]} should be accessible via NumericCellValue");
                }
            }

            // Test with XSSF (XLSX format)
            using (var xssfWorkbook = new XSSFWorkbook())
            {
                var xssfSheet = xssfWorkbook.CreateSheet("Test");

                for (int i = 0; i < testValues.Length; i++)
                {
                    var xssfRow = xssfSheet.CreateRow(i);
                    var xssfCell = xssfRow.CreateCell(0);

                    // Set numeric value without date formatting
                    xssfCell.SetCellValue(testValues[i]);

                    // Verify it's NOT date formatted
                    Assert.IsFalse(DateUtil.IsCellDateFormatted(xssfCell),
                        $"Value {testValues[i]} should not be date formatted");

                    // DateCellValue should return null
                    Assert.IsNull(xssfCell.DateCellValue,
                        $"Numeric value {testValues[i]} should return null for DateCellValue");

                    // But numeric access should still work
                    Assert.AreEqual(testValues[i], xssfCell.NumericCellValue, 0.001,
                        $"Numeric value {testValues[i]} should be accessible via NumericCellValue");
                }
            }
        }

        /// <summary>
        /// Test that manually applying date formatting to numeric cells makes DateCellValue work
        /// This verifies the IsCellDateFormatted check works correctly
        /// </summary>
        [Test]
        public void TestManuallyFormattedDateCells_ShouldReturnDateCellValue()
        {
            // Excel date serial number for 2023-12-25 (days since 1900-01-01)
            var expectedDate = new DateTime(2023, 12, 25);
            double dateSerial = DateUtil.GetExcelDate(expectedDate, false); // Calculate correct serial number

            // Test with HSSF (XLS format)
            using (var hssfWorkbook = new HSSFWorkbook())
            {
                var hssfSheet = hssfWorkbook.CreateSheet("Test");
                var hssfRow = hssfSheet.CreateRow(0);
                var hssfCell = hssfRow.CreateCell(0);

                // Set numeric value first
                hssfCell.SetCellValue(dateSerial);
                Assert.IsNull(hssfCell.DateCellValue, "Should be null before date formatting");

                // Apply date formatting
                var dateStyle = hssfWorkbook.CreateCellStyle();
                dateStyle.DataFormat = hssfWorkbook.CreateDataFormat().GetFormat("mm/dd/yyyy");
                hssfCell.CellStyle = dateStyle;

                // Now it should be recognized as a date
                Assert.IsTrue(DateUtil.IsCellDateFormatted(hssfCell));
                Assert.IsNotNull(hssfCell.DateCellValue);
                Assert.AreEqual(expectedDate.Date, hssfCell.DateCellValue.Value.Date);
            }

            // Test with XSSF (XLSX format)
            using (var xssfWorkbook = new XSSFWorkbook())
            {
                var xssfSheet = xssfWorkbook.CreateSheet("Test");
                var xssfRow = xssfSheet.CreateRow(0);
                var xssfCell = xssfRow.CreateCell(0);

                // Set numeric value first
                xssfCell.SetCellValue(dateSerial);
                Assert.IsNull(xssfCell.DateCellValue, "Should be null before date formatting");

                // Apply date formatting
                var dateStyle = xssfWorkbook.CreateCellStyle();
                dateStyle.DataFormat = xssfWorkbook.CreateDataFormat().GetFormat("mm/dd/yyyy");
                xssfCell.CellStyle = dateStyle;

                // Now it should be recognized as a date
                Assert.IsTrue(DateUtil.IsCellDateFormatted(xssfCell));
                Assert.IsNotNull(xssfCell.DateCellValue);
                Assert.AreEqual(expectedDate.Date, xssfCell.DateCellValue.Value.Date);
            }
        }
    }
}
