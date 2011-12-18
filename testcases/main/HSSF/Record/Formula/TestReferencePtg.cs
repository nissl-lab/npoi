/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record.Formula
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.UserModel;

    using TestCases.HSSF;
    using NPOI.HSSF.Record;
    using NPOI.Util.IO;

    /**
     * Tests for {@link RefPtg}.
     */
    [TestClass]
    public class TestReferencePtg
    {
        /**
         * Tests Reading a file containing this ptg.
         */
        [TestMethod]
        public void TestReading()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook("ReferencePtg.xls");
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);

            // First row
            Assert.AreEqual(55.0,
                         sheet.GetRow(0).GetCell((short)0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(55.0,
                         sheet.GetRow(0).GetCell((short)1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A1",
                         sheet.GetRow(0).GetCell((short)1).CellFormula, "Wrong formula string for reference");

            // Now moving over the 2**15 boundary
            // (Remember that excel row (n) is1 poi row (n-1)
            Assert.AreEqual(32767.0,
                    sheet.GetRow(32766).GetCell((short)0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(32767.0,
                    sheet.GetRow(32766).GetCell((short)1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A32767",
                    sheet.GetRow(32766).GetCell((short)1).CellFormula, "Wrong formula string for reference");

            Assert.AreEqual(32768.0,
                    sheet.GetRow(32767).GetCell((short)0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(32768.0,
                    sheet.GetRow(32767).GetCell((short)1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A32768",
                    sheet.GetRow(32767).GetCell((short)1).CellFormula, "Wrong formula string for reference");

            Assert.AreEqual(32769.0,
                    sheet.GetRow(32768).GetCell((short)0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(32769.0,
                    sheet.GetRow(32768).GetCell((short)1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A32769",
                    sheet.GetRow(32768).GetCell((short)1).CellFormula, "Wrong formula string for reference");

            Assert.AreEqual(32770.0,
                    sheet.GetRow(32769).GetCell((short)0).NumericCellValue, 0.0, "Wrong numeric value for original number");
            Assert.AreEqual(32770.0,
                    sheet.GetRow(32769).GetCell((short)1).NumericCellValue, 0.0, "Wrong numeric value for referemce");
            Assert.AreEqual("A32770",
                    sheet.GetRow(32769).GetCell((short)1).CellFormula, "Wrong formula string for reference");
        }
        [TestMethod]
        public void TestBug44921()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex44921-21902.xls");

            try
            {
                HSSFTestDataSamples.WriteOutAndReadBack(wb);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Coding Error: This method should never be called. This ptg should be converted"))
                {
                    throw new AssertFailedException("Identified bug 44921");
                }
                throw e;
            }
        }
        private static byte[] tRefN_data = {
    	0x2C, 33, 44, 55, 66,
        };
        [TestMethod]
        public void TestReadWrite_tRefN_bug45091()
        {
            LittleEndianInput in1 = TestcaseRecordInputStream.CreateLittleEndian(tRefN_data);
            Ptg[] ptgs = Ptg.ReadTokens(tRefN_data.Length, in1);
            byte[] outData = new byte[5];
            Ptg.SerializePtgs(ptgs, outData, 0);
            if (outData[0] == 0x24)
            {
                throw new AssertFailedException("Identified bug 45091");
            }
            Assert.IsTrue(NPOI.Util.Arrays.Equals(tRefN_data, outData));
        }
    }
}
