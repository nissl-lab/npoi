/*
 *  ====================================================================
 *    Licensed to the collaborators of the NPOI project under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The collaborators license this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;

namespace TestCases.SS.Formula.Functions
{
    /// <summary>
    /// This class provides a simple comparison test
    /// to show if the cached formula result in an Excel file
    /// is still same after being recalculated by NPOI.
    /// The specific class is for numeric and error values.
    /// </summary>
    [TestFixture]
    public abstract class CompareNumericFuncEvalTestBase
    {
        public abstract string TestFileName { get; }
        // In real-world Excel's tolerance control is more complicated.
        // Save it for now.
        private const double Tolerance = 1e-7;
        private XSSFWorkbook _workbook;
        [OneTimeSetUp]
        public void LoadData()
        {
            var fldr = Path.Combine(TestContext.CurrentContext.TestDirectory, TestContext.Parameters["function"]);
            var file = Path.Combine(fldr, TestFileName);

            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                _workbook = new XSSFWorkbook(fs);
            }
        }

        [OneTimeTearDown]
        public void Dispose()
        {
            _workbook?.Close();
        }

        internal static SortedList<string, double> GetSpreadsheetContent(IWorkbook workbook)
        {
            var list = new SortedList<string, double>();

            var sheet = workbook.GetSheetAt(0);

            for (var rowId = sheet.FirstRowNum; rowId <= sheet.LastRowNum; rowId++)
            {
                var row = sheet.GetRow(rowId);
                if (row is null || row.FirstCellNum < 0)
                    continue;

                for (var colId = row.FirstCellNum; colId <= row.LastCellNum; colId++)
                {
                    var cell = row.GetCell(colId);
                    if (cell is null) continue;
                    if (cell.CellType != CellType.Formula) continue;
                    if (cell.CachedFormulaResultType != CellType.Numeric
                        && cell.CachedFormulaResultType != CellType.Error) continue;

                    list[cell.Address.FormatAsString()] =
                        cell.CachedFormulaResultType == CellType.Numeric ?
                        cell.NumericCellValue : cell.ErrorCellValue;
                }
            }
            return list;
        }

        [Test]
        public void TestEvaluate()
        {
            var originalData = GetSpreadsheetContent(_workbook);

            var evaluator = new XSSFFormulaEvaluator(_workbook);
            evaluator.ClearAllCachedResultValues();
            Assert.DoesNotThrow(() => evaluator.EvaluateAll());

            var evaluatedData = GetSpreadsheetContent(_workbook);

            Assert.Multiple(() =>
            {
                foreach (var kv in evaluatedData)
                {
                    if (!originalData.TryGetValue(kv.Key, out var val))
                    {
                        Assert.Fail($"Spreadsheet structure changed! No {kv.Key} cell in the original spreadsheet.");
                        break;
                    }

                    Assert.AreEqual(val, kv.Value, Tolerance, kv.Key);
                }
            });
        }
    }
}
