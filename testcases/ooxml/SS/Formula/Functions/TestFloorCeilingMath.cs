/*
 *  ====================================================================
 *    Licensed to the collaborators of the NPOI project under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The collaborators licenses this file to You under the Apache License, Version 2.0
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
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NUnit.Framework.Constraints;
using NPOI.SS.Util;

namespace TestCases.SS.Formula.Functions
{
    /// <summary>
    /// Testing FLOOR.MATH & CEILING.MATH
    /// </summary>
    [TestFixture]
    public class TestFloorCeilingMath
    {
        // In real-world Excel's tolerance control is more complicated.
        // Save it for now.
        private const double Tolerance = 1e-7;
        public enum FunctionTested
        {
            Ceiling,
            Floor
        }
        public sealed class TestFunction
        {
            public TestFunction(FunctionTested func, bool mode)
            {
                Function = func;
                Mode = mode;
            }
            public FunctionTested Function { get; set; }
            public bool Mode { get; set; }
            public string SheetName
                => (Function == FunctionTested.Ceiling ? "CEILING" : "FLOOR") + "," + Mode.ToString().ToUpperInvariant();

            public FloorCeilingMathBase Evaluator 
                => Function == FunctionTested.Ceiling ? CeilingMath.Instance : (FloorCeilingMathBase)FloorMath.Instance;

            public double Evaluate(double number, double significance)
                => Evaluator.Evaluate(number, significance, Mode);
            public override string ToString()
                => SheetName;
        }
        
        private XSSFWorkbook _workbook;
        [OneTimeSetUp]
        public void LoadData()
        {
            var fldr = Path.Combine(TestContext.CurrentContext.TestDirectory, TestContext.Parameters["function"]);
            const string filename = "FloorCeilingMath.xlsx";
            var file = Path.Combine(fldr, filename);

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
        public static TestFunction[] FunctionVariants => new[]
        {
            new TestFunction(FunctionTested.Ceiling, false),
            new TestFunction(FunctionTested.Ceiling, true),
            new TestFunction(FunctionTested.Floor, false),
            new TestFunction(FunctionTested.Floor, true),
        };

        [Test, Order(1)]
        [TestCaseSource(nameof(FunctionVariants))]
        public void TestEvaluate(TestFunction function)
        {
            const int StartRowIndex = 1;
            const int StartColumnIndex = 0;
            const int Count = 34;

            Assert.Multiple(() =>
            {
                var sheet = _workbook.GetSheet(function.SheetName);
                for (var i = 1; i <= Count; i++)
                {
                    var row = sheet.GetRow(i + StartRowIndex);
                    var significance = row.GetCell(StartColumnIndex).NumericCellValue;

                    for (var j = 1; j <= Count; j++)
                    {
                        var number = sheet.GetRow(StartRowIndex).GetCell(j + StartColumnIndex).NumericCellValue;
                        var expected = row.GetCell(j + StartColumnIndex).NumericCellValue;

                        var functionResult = function.Evaluate(number, significance);
                        
                        Assert.AreEqual(expected, functionResult, Tolerance, $"{function}, {number}, {significance}");
                    }
                }
            });
        }
        [Test, Order(2), NonParallelizable]
        public void EvaluateAllFormulas()
        {
            var evaluator = new XSSFFormulaEvaluator(_workbook);
            evaluator.ClearAllCachedResultValues();
            Assert.DoesNotThrow(() => evaluator.EvaluateAll());
        }
    }
}
