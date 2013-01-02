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

    using NUnit.Framework;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record.Common;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;
    using NPOI.HSSF.Record;
    using TestCases.HSSF.UserModel;
    /**
     * Tests for <tt>FeatRecord</tt>
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestFeatRecord
    {
        public void TestWithoutFeatRecord()
        {
            HSSFWorkbook hssf =
                   HSSFTestDataSamples.OpenSampleWorkbook("46136-WithWarnings.xls");
            InternalWorkbook wb = HSSFTestHelper.GetWorkbookForTest(hssf);

            Assert.AreEqual(1, hssf.NumberOfSheets);

            int countFR = 0;
            int countFRH = 0;

            // Check on the workbook, but shouldn't be there!
            foreach (Record r in wb.Records)
            {
                if (r is FeatRecord)
                {
                    countFR++;
                }
                else if (r.Sid == FeatRecord.sid)
                {
                    countFR++;
                }
                if (r is FeatHdrRecord)
                {
                    countFRH++;
                }
                else if (r.Sid == FeatHdrRecord.sid)
                {
                    countFRH++;
                }
            }

            Assert.AreEqual(0, countFR);
            Assert.AreEqual(0, countFRH);

            // Now check on the sheet
            HSSFSheet s = (HSSFSheet)hssf.GetSheetAt(0);
            InternalSheet sheet = HSSFTestHelper.GetSheetForTest(s);

            foreach (RecordBase rb in sheet.Records)
            {
                if (rb is Record)
                {
                    Record r = (Record)rb;
                    if (r is FeatRecord)
                    {
                        countFR++;
                    }
                    else if (r.Sid == FeatRecord.sid)
                    {
                        countFR++;
                    }
                    if (r is FeatHdrRecord)
                    {
                        countFRH++;
                    }
                    else if (r.Sid == FeatHdrRecord.sid)
                    {
                        countFRH++;
                    }
                }
            }

            Assert.AreEqual(0, countFR);
            Assert.AreEqual(0, countFRH);
        }

        public void TestReadFeatRecord()
        {
            HSSFWorkbook hssf =
                   HSSFTestDataSamples.OpenSampleWorkbook("46136-NoWarnings.xls");
            InternalWorkbook wb = HSSFTestHelper.GetWorkbookForTest(hssf);

            FeatRecord fr = null;
            FeatHdrRecord fhr = null;

            Assert.AreEqual(1, hssf.NumberOfSheets);

            // First check it isn't on the Workbook
            int countFR = 0;
            int countFRH = 0;
            foreach (Record r in wb.Records)
            {
                if (r is FeatRecord)
                {
                    fr = (FeatRecord)r;
                    countFR++;
                }
                else if (r.Sid == FeatRecord.sid)
                {
                    Assert.Fail("FeatRecord SID found but not Created correctly!");
                }
                if (r is FeatHdrRecord)
                {
                    countFRH++;
                }
                else if (r.Sid == FeatHdrRecord.sid)
                {
                    Assert.Fail("FeatHdrRecord SID found but not Created correctly!");
                }
            }

            Assert.AreEqual(0, countFR);
            Assert.AreEqual(0, countFRH);

            // Now find it on our sheet
            HSSFSheet s = (HSSFSheet)hssf.GetSheetAt(0);
            InternalSheet sheet = HSSFTestHelper.GetSheetForTest(s);

            foreach (RecordBase rb in sheet.Records)
            {
                if (rb is Record)
                {
                    Record r = (Record)rb;
                    if (r is FeatRecord)
                    {
                        fr = (FeatRecord)r;
                        countFR++;
                    }
                    else if (r.Sid == FeatRecord.sid)
                    {
                        countFR++;
                    }
                    if (r is FeatHdrRecord)
                    {
                        fhr = (FeatHdrRecord)r;
                        countFRH++;
                    }
                    else if (r.Sid == FeatHdrRecord.sid)
                    {
                        countFRH++;
                    }
                }
            }

            Assert.AreEqual(1, countFR);
            Assert.AreEqual(1, countFRH);
            Assert.IsNotNull(fr);
            Assert.IsNotNull(fhr);

            // Now check the contents are as expected
            Assert.AreEqual(
                    FeatHdrRecord.SHAREDFEATURES_ISFFEC2,
                    fr.Isf_sharedFeatureType
            );

            // Applies to one cell only
            Assert.AreEqual(1, fr.CellRefs.Length);
            Assert.AreEqual(0, fr.CellRefs[0].FirstRow);
            Assert.AreEqual(0, fr.CellRefs[0].LastRow);
            Assert.AreEqual(0, fr.CellRefs[0].FirstColumn);
            Assert.AreEqual(0, fr.CellRefs[0].LastColumn);

            // More Checking of shared features stuff
            Assert.AreEqual(4, fr.CbFeatData);
            Assert.AreEqual(4, fr.SharedFeature.DataSize);
            Assert.AreEqual(typeof(FeatFormulaErr2), fr.SharedFeature.GetType());

            FeatFormulaErr2 fferr2 = (FeatFormulaErr2)fr.SharedFeature;
            Assert.AreEqual(0x04, fferr2.RawErrorCheckValue);

            Assert.IsFalse(fferr2.CheckCalculationErrors);
            Assert.IsFalse(fferr2.CheckDateTimeFormats);
            Assert.IsFalse(fferr2.CheckEmptyCellRef);
            Assert.IsFalse(fferr2.CheckInconsistentFormulas);
            Assert.IsFalse(fferr2.CheckInconsistentRanges);
            Assert.IsTrue(fferr2.CheckNumbersAsText);
            Assert.IsFalse(fferr2.CheckUnprotectedFormulas);
            Assert.IsFalse(fferr2.PerformDataValidation);
        }

        /**
         *  cloning sheets with feat records 
         */
        public void TestCloneSheetWithFeatRecord()
        {
            IWorkbook wb =
                HSSFTestDataSamples.OpenSampleWorkbook("46136-WithWarnings.xls");
            wb.CloneSheet(0);
        }
    }

}