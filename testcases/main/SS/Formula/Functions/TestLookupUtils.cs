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
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.UDF;
    using NPOI.SS.UserModel;

    /**
     * Tests that whole-column references (e.g. A:A) used by lookup functions
     * are backed by a vector sized to the sheet's actual used range rather than
     * the full 65536/1048576 row height of the reference itself. Without the
     * clamp, every lookup against a whole column walks the entire row range of
     * the sheet version even when almost all of it is blank.
     */
    [TestFixture]
    public class TestLookupUtils
    {
        // captures the AreaEval a formula passes to it and builds a ColumnVector
        // from it the same way VLOOKUP/XLOOKUP do internally
        private class CaptureColumnVectorSizeFunc : FreeRefFunction
        {
            public int? CapturedSize;

            public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
            {
                TwoDEval area = (TwoDEval)args[0];
                CapturedSize = LookupUtils.CreateColumnVector(area, 0).Size;
                return NumberEval.ZERO;
            }
        }

        [Test]
        public void ColumnVectorForWholeColumnRefIsClampedToSheetUsedRange()
        {
            using (HSSFWorkbook wb = new HSSFWorkbook())
            {
                ISheet dataSheet = wb.CreateSheet("Data");
                for (int r = 0; r <= 4; r++)
                {
                    dataSheet.CreateRow(r).CreateCell(0).SetCellValue(r);
                }
                ClassicAssert.AreEqual(4, dataSheet.LastRowNum);

                CaptureColumnVectorSizeFunc capture = new CaptureColumnVectorSizeFunc();
                wb.AddToolPack(new DefaultUDFFinder(
                        new String[] { "CAPTURESIZE" },
                        new FreeRefFunction[] { capture }));

                ISheet calcSheet = wb.CreateSheet("Calc");
                ICell cell = calcSheet.CreateRow(0).CreateCell(0);
                cell.SetCellFormula("CAPTURESIZE(Data!A:A)");

                IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();
                fe.EvaluateFormulaCell(cell);

                ClassicAssert.IsNotNull(capture.CapturedSize, "custom function was never invoked");
                // A:A spans rows 0..65535 in the BIFF8 (xls) sheet version, but the
                // sheet only has 5 rows of real data - the lookup vector must not
                // scan the other ~65531 guaranteed-blank rows.
                ClassicAssert.AreEqual(dataSheet.LastRowNum + 1, capture.CapturedSize.Value,
                        "ColumnVector should be clamped to the sheet's used row range, not the whole-column height");
            }
        }

        [Test]
        public void XLookupOverWholeColumnStillFindsValuesWithinUsedRange()
        {
            using (HSSFWorkbook wb = new HSSFWorkbook())
            {
                ISheet dataSheet = wb.CreateSheet("Data");
                SS.Util.Utils.AddRow(dataSheet, 0, "k0", "v0");
                SS.Util.Utils.AddRow(dataSheet, 1, "k1", "v1");
                SS.Util.Utils.AddRow(dataSheet, 2, "k2", "v2");

                ISheet calcSheet = wb.CreateSheet("Calc");
                ICell cell = calcSheet.CreateRow(0).CreateCell(0);

                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                SS.Util.Utils.AssertString(fe, cell, "_xlfn.XLOOKUP(\"k1\",Data!A:A,Data!B:B,\"missing\")", "v1");
                SS.Util.Utils.AssertString(fe, cell, "_xlfn.XLOOKUP(\"not-there\",Data!A:A,Data!B:B,\"missing\")", "missing");
            }
        }
    }
}
