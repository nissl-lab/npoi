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

namespace TestCases.SS.Formula.PTG
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    using TestCases.HSSF;
    using TestCases.HSSF.Record;

    /**
     * Tests for {@link RefPtg}.
     */
    [TestFixture]
    public class TestReferencePtg
    {
        /**
         * Tests Reading a file Containing this ptg.
         */
        [Test]
        public void TestReading()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("ReferencePtg.xls");
            ISheet sheet = workbook.GetSheetAt(0);

            // First row
            Assert.AreEqual(55.0, sheet.GetRow(0).GetCell(0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(55.0, sheet.GetRow(0).GetCell(1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A1", sheet.GetRow(0).GetCell(1).CellFormula, "Wrong formula string for reference");

            // Now moving over the 2**15 boundary
            // (Remember that excel row (n) is poi row (n-1)
            Assert.AreEqual(32767.0, sheet.GetRow(32766).GetCell(0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(32767.0, sheet.GetRow(32766).GetCell(1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A32767", sheet.GetRow(32766).GetCell(1).CellFormula, "Wrong formula string for reference");

            Assert.AreEqual(32768.0, sheet.GetRow(32767).GetCell(0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(32768.0, sheet.GetRow(32767).GetCell(1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A32768", sheet.GetRow(32767).GetCell(1).CellFormula, "Wrong formula string for reference");

            Assert.AreEqual(32769.0, sheet.GetRow(32768).GetCell(0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(32769.0, sheet.GetRow(32768).GetCell(1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A32769", sheet.GetRow(32768).GetCell(1).CellFormula, "Wrong formula string for reference");

            Assert.AreEqual(32770.0, sheet.GetRow(32769).GetCell(0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(32770.0, sheet.GetRow(32769).GetCell(1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A32770", sheet.GetRow(32769).GetCell(1).CellFormula, "Wrong formula string for reference");
        }
        [Test]
        public void TestBug44921()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex44921-21902.xls");

            try
            {
                HSSFTestDataSamples.WriteOutAndReadBack(wb);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Coding Error: This method should never be called. This ptg should be Converted"))
                {
                    throw new AssertionException("Identified bug 44921");
                }
                throw;
            }
        }
        private static byte[] tRefN_data = {
        0x2C, 33, 44, 55, 66,
    };
        [Test]
        public void TestReadWrite_tRefN_bug45091()
        {
            ILittleEndianInput in1 = TestcaseRecordInputStream.CreateLittleEndian(tRefN_data);
            Ptg[] ptgs = Ptg.ReadTokens(tRefN_data.Length, in1);
            byte[] outData = new byte[5];
            Ptg.SerializePtgs(ptgs, outData, 0);
            if (outData[0] == 0x24)
            {
                throw new AssertionException("Identified bug 45091");
            }
            Assert.IsTrue(Arrays.Equals(tRefN_data, outData));
        }

        /**
         * Test that RefPtgBase can handle references with column index greater than 255,
         * see Bugzilla 50096
         */
        [Test]
        public void TestColumnGreater255()
        {
            RefPtgBase ptg;
            ptg = new RefPtg("IW1");
            Assert.AreEqual(256, ptg.Column);
            Assert.AreEqual("IW1", ptg.FormatReferenceAsString());

            ptg = new RefPtg("JA1");
            Assert.AreEqual(260, ptg.Column);
            Assert.AreEqual("JA1", ptg.FormatReferenceAsString());
        }
    }


}