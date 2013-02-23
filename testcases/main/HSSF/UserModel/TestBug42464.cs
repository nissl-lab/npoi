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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections;
    using System.IO;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS.Formula;
    using NPOI.HSSF.UserModel;
    
    using NUnit.Framework;

    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.SS.Formula.PTG;

    /**
     * 
     */
    [TestFixture]
    public class TestBug42464
    {
        [Test]
        public void TestOKFile()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("42464-ExpPtg-ok.xls");
            Process(wb);
        }
        [Test]
        public void TestExpSharedBadFile()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("42464-ExpPtg-bad.xls");
            Process(wb);
        }

        private static void Process(HSSFWorkbook wb)
        {
            HSSFFormulaEvaluator eval = new HSSFFormulaEvaluator(wb);
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                NPOI.SS.UserModel.ISheet s = wb.GetSheetAt(i);

                IEnumerator it = s.GetRowEnumerator();
                while (it.MoveNext())
                {
                    IRow r = (IRow)it.Current;
                    Process(r, eval);
                }
            }
        }

        private static void Process(IRow row, HSSFFormulaEvaluator eval)
        {
            IEnumerator it = row.GetEnumerator();
            while (it.MoveNext())
            {
                ICell cell = (ICell)it.Current;
                if (cell.CellType != NPOI.SS.UserModel.CellType.Formula)
                {
                    continue;
                }
                FormulaRecordAggregate record = (FormulaRecordAggregate)((HSSFCell)cell).CellValueRecord;
                FormulaRecord r = record.FormulaRecord;
                Ptg[] ptgs = r.ParsedExpression;

                String cellRef = new CellReference(row.RowNum, cell.ColumnIndex, false, false).FormatAsString();
#if !HIDE_UNREACHABLE_CODE
                if (false && cellRef.Equals("BP24"))
                { 
                    Console.Write(cellRef);
                    Console.WriteLine(" - has " + ptgs.Length + " ptgs:");
                    for (int i = 0; i < ptgs.Length; i++)
                    {
                        String c = ptgs[i].GetType().ToString();
                        Console.WriteLine("\t" + c.Substring(c.LastIndexOf('.') + 1));
                    }
                    Console.WriteLine("-> " + cell.CellFormula);
                }
#endif

                NPOI.SS.UserModel.CellValue evalResult = eval.Evaluate(cell);
                Assert.IsNotNull(evalResult);
            }
        }
    }
}