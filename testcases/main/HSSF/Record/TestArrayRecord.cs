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

namespace TestCases.HSSF.Record
{
    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Util;
    using NPOI.Util;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;

    [TestFixture]
    public class TestArrayRecord
    {
        [Test]
        public void TestRead()
        {
            String hex =
                    "21 02 25 00 01 00 01 00 01 01 00 00 00 00 00 00 " +
                    "17 00 65 00 00 01 00 02 C0 02 C0 65 00 00 01 00 " +
                    "03 C0 03 C0 04 62 01 07 00";
            byte[] data = HexRead.ReadFromString(hex);
            RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
            ArrayRecord r1 = new ArrayRecord(in1);
            CellRangeAddress8Bit range = r1.Range;
            Assert.AreEqual(1, range.FirstColumn);
            Assert.AreEqual(1, range.LastColumn);
            Assert.AreEqual(1, range.FirstRow);
            Assert.AreEqual(1, range.LastRow);

            Ptg[] ptg = r1.FormulaTokens;
            Assert.AreEqual(FormulaRenderer.ToFormulaString(null, ptg), "MAX(C1:C2-D1:D2)");

            //construct a new ArrayRecord with the same contents as r1
            Ptg[] fmlaPtg = FormulaParser.Parse("MAX(C1:C2-D1:D2)", null, FormulaType.Array, 0);
            ArrayRecord r2 = new ArrayRecord(Formula.Create(fmlaPtg), new CellRangeAddress8Bit(1, 1, 1, 1));
            byte[] ser = r2.Serialize();
            //serialize and check that the data is the same as in r1
            Assert.AreEqual(HexDump.ToHex(data), HexDump.ToHex(ser));


        }

        [Test]
        public void TestBug57231()
        {
            HSSFWorkbook wb = HSSFTestDataSamples
                    .OpenSampleWorkbook("57231_MixedGasReport.xls");
            HSSFSheet sheet = wb.GetSheet("master") as HSSFSheet;

            HSSFSheet newSheet = wb.CloneSheet(wb.GetSheetIndex(sheet)) as HSSFSheet;
            int idx = wb.GetSheetIndex(newSheet);
            wb.SetSheetName(idx, "newName");

            // Write the output to a file
            HSSFWorkbook wbBack = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);

            Assert.IsNotNull(wbBack.GetSheet("master"));
            Assert.IsNotNull(wbBack.GetSheet("newName"));
        }

    }
}