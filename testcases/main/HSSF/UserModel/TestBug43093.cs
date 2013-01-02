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
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    using NPOI.SS.UserModel;
    /**
     * 
     */
    [TestFixture]
    public class TestBug43093Class
    {

        private static void AddNewSheetWithCellsA1toD4(HSSFWorkbook book, int sheet)
        {

            NPOI.SS.UserModel.ISheet sht = book.CreateSheet("s" + sheet);
            for (int r = 0; r < 4; r++)
            {

                IRow row = sht.CreateRow(r);
                for (int c = 0; c < 4; c++)
                {

                    ICell cel = row.CreateCell(c);
                    cel.SetCellValue(sheet * 100 + r * 10 + c);
                }
            }
        }
        [Test]
        public void TestBug43093()
        {
            HSSFWorkbook xlw = new HSSFWorkbook();

            AddNewSheetWithCellsA1toD4(xlw, 1);
            AddNewSheetWithCellsA1toD4(xlw, 2);
            AddNewSheetWithCellsA1toD4(xlw, 3);
            AddNewSheetWithCellsA1toD4(xlw, 4);

            NPOI.SS.UserModel.ISheet s2 = xlw.GetSheet("s2");
            IRow s2r3 = s2.GetRow(3);
            ICell s2E4 = s2r3.CreateCell(4);
            s2E4.CellFormula = ("SUM(s3!B2:C3)");

            HSSFFormulaEvaluator eva = new HSSFFormulaEvaluator(xlw);
            double d = eva.Evaluate(s2E4).NumberValue;

            // internalEvaluate(...) Area3DEval.: 311+312+321+322 expected
            Assert.AreEqual(d, (311 + 312 + 321 + 322), 0.0000001);
            // System.out.println("Area3DEval ok.: 311+312+321+322=" + d);
        }
    }
}
