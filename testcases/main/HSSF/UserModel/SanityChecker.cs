
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


namespace TestCases.HSSF.UserModel
{
    using System;
    using System.Collections;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.HSSF.UserModel;

    using NUnit.Framework;

    /**
     * Designed to Check wither the records written actually make sense.
     */
    public class SanityChecker
    {
        public class CheckRecord
        {
            Type record;
            char occurance;  // 1 = one time, M = 1..many times, * = 0..many, 0 = optional
            private bool together;

            public CheckRecord(Type record, char occurance)
                : this(record, occurance, true)
            {

            }

            /**
             * @param record        The record type to Check
             * @param occurance     The occurance 1 = occurs once, M = occurs many times
             * @param together
             */
            public CheckRecord(Type record, char occurance, bool together)
            {
                this.record = record;
                this.occurance = occurance;
                this.together = together;
            }

            public Type GetRecord()
            {
                return record;
            }

            public char GetOccurance()
            {
                return occurance;
            }

            public bool IsRequired()
            {
                return occurance == '1' || occurance == 'M';
            }

            public bool IsOptional()
            {
                return occurance == '0' || occurance == '*';
            }

            public bool Istogether()
            {
                return together;
            }

            public bool IsMany()
            {
                return occurance == '*' || occurance == 'M';
            }

            public int Match(IList records, int recordIdx)
            {
                int firstRecord = FindFirstRecord(records, GetRecord(), recordIdx);
                if (IsRequired())
                {
                    return MatchRequired(firstRecord, records, recordIdx);
                }
                else
                {
                    return MatchOptional(firstRecord, records, recordIdx);
                }
            }

            private int MatchOptional(int firstRecord, IList records, int recordIdx)
            {
                if (firstRecord == -1)
                {
                    return recordIdx;
                }

                return MatchOneOrMany(records, firstRecord);
            }

            private int MatchRequired(int firstRecord, IList records, int recordIdx)
            {
                if (firstRecord == -1)
                {
                    Assert.Fail("Manditory record missing or out of order: " + record);
                }

                return MatchOneOrMany(records, firstRecord);
            }

            private int MatchOneOrMany(IList records, int recordIdx)
            {
                if (IsZeroOrOne())
                {
                    // Check no other records
                    if (FindFirstRecord(records, GetRecord(), recordIdx + 1) != -1)
                        Assert.Fail("More than one record Matched for " + GetRecord().Name);
                }
                else if (IsZeroToMany())
                {
                    if (together)
                    {
                        int nextIdx = FindFirstRecord(records, record, recordIdx + 1);
                        while (nextIdx != -1)
                        {
                            if (nextIdx - 1 != recordIdx)
                                Assert.Fail("Records are not together " + record.Name);
                            recordIdx = nextIdx;
                            nextIdx = FindFirstRecord(records, record, recordIdx + 1);
                        }
                    }
                }
                return recordIdx + 1;
            }

            private bool IsZeroToMany()
            {
                return occurance == '*' || occurance == 'M';
            }

            private bool IsZeroOrOne()
            {
                return occurance == '0' || occurance == '1';
            }
        }

        CheckRecord[] workbookRecords = new CheckRecord[] {
        new CheckRecord(typeof(BOFRecord), '1'),
        new CheckRecord(typeof(InterfaceHdrRecord), '1'),
        new CheckRecord(typeof(MMSRecord), '1'),
        new CheckRecord(typeof(InterfaceEndRecord), '1'),
        new CheckRecord(typeof(WriteAccessRecord), '1'),
        new CheckRecord(typeof(CodepageRecord), '1'),
        new CheckRecord(typeof(DSFRecord), '1'),
        new CheckRecord(typeof(TabIdRecord), '1'),
        new CheckRecord(typeof(FnGroupCountRecord), '1'),
        new CheckRecord(typeof(WindowProtectRecord), '1'),
        new CheckRecord(typeof(ProtectRecord), '1'),
        new CheckRecord(typeof(PasswordRev4Record), '1'),
        new CheckRecord(typeof(WindowOneRecord), '1'),
        new CheckRecord(typeof(BackupRecord), '1'),
        new CheckRecord(typeof(HideObjRecord), '1'),
        new CheckRecord(typeof(DateWindow1904Record), '1'),
        new CheckRecord(typeof(PrecisionRecord), '1'),
        new CheckRecord(typeof(RefreshAllRecord), '1'),
        new CheckRecord(typeof(BookBoolRecord), '1'),
        new CheckRecord(typeof(FontRecord), 'M'),
        new CheckRecord(typeof(FormatRecord), 'M'),
        new CheckRecord(typeof(ExtendedFormatRecord), 'M'),
        new CheckRecord(typeof(StyleRecord), 'M'),
        new CheckRecord(typeof(UseSelFSRecord), '1'),
        new CheckRecord(typeof(BoundSheetRecord), 'M'),
        new CheckRecord(typeof(CountryRecord), '1'),
        new CheckRecord(typeof(SupBookRecord), '0'),
        new CheckRecord(typeof(ExternSheetRecord), '0'),
        new CheckRecord(typeof(NameRecord), '*'),
        new CheckRecord(typeof(SSTRecord), '1'),
        new CheckRecord(typeof(ExtSSTRecord), '1'),
        new CheckRecord(typeof(EOFRecord), '1'),
    };

        CheckRecord[] sheetRecords = new CheckRecord[] {
        new CheckRecord(typeof(BOFRecord), '1'),
        new CheckRecord(typeof(CalcModeRecord), '1'),
        new CheckRecord(typeof(RefModeRecord), '1'),
        new CheckRecord(typeof(IterationRecord), '1'),
        new CheckRecord(typeof(DeltaRecord), '1'),
        new CheckRecord(typeof(SaveRecalcRecord), '1'),
        new CheckRecord(typeof(PrintHeadersRecord), '1'),
        new CheckRecord(typeof(PrintGridlinesRecord), '1'),
        new CheckRecord(typeof(GridsetRecord), '1'),
        new CheckRecord(typeof(GutsRecord), '1'),
        new CheckRecord(typeof(DefaultRowHeightRecord), '1'),
        new CheckRecord(typeof(WSBoolRecord), '1'),
        new CheckRecord(typeof(PageSettingsBlock), '1'),
        new CheckRecord(typeof(DefaultColWidthRecord), '1'),
        new CheckRecord(typeof(DimensionsRecord), '1'),
        new CheckRecord(typeof(WindowTwoRecord), '1'),
        new CheckRecord(typeof(SelectionRecord), '1'),
        new CheckRecord(typeof(EOFRecord), '1')
    };

        private void CheckWorkbookRecords(InternalWorkbook workbook)
        {
            IList records = workbook.Records;
            Assert.IsTrue(records[0] is BOFRecord);
            Assert.IsTrue(records[records.Count - 1] is EOFRecord);

            CheckRecordOrder(records, workbookRecords);
            //        CheckRecordstogether(records, workbookRecords);
        }

        private void CheckSheetRecords(InternalSheet sheet)
        {
            IList records = sheet.Records;
            Assert.IsTrue(records[0] is BOFRecord);
            Assert.IsTrue(records[records.Count - 1] is EOFRecord);

            CheckRecordOrder(records, sheetRecords);
            //        CheckRecordstogether(records, sheetRecords);
        }

        public void CheckHSSFWorkbook(HSSFWorkbook wb)
        {
            CheckWorkbookRecords(wb.Workbook);
            for (int i = 0; i < wb.NumberOfSheets; i++)
                CheckSheetRecords(((HSSFSheet)wb.GetSheetAt(i)).Sheet);

        }

        /*
        private void CheckRecordstogether(List records, CheckRecord[] Check)
        {
            for ( int CheckIdx = 0; CheckIdx < Check.Length; CheckIdx++ )
            {
                int recordIdx = FindFirstRecord(records, Check[CheckIdx].GetRecord());
                bool notFoundAndRecordRequired = (recordIdx == -1 && Check[CheckIdx].isRequired());
                if (notFoundAndRecordRequired)
                {
                    Assert.Fail("Expected to Find record of class " + Check.GetClass() + " but did not");
                }
                else if (recordIdx >= 0)
                {
                    if (Check[CheckIdx].isMany())
                    {
                        // Skip records that are together
                        while (recordIdx < records.Count && Check[CheckIdx].GetRecord().isInstance(records.Get(recordIdx)))
                            recordIdx++;
                    }

                    // Make sure record does not occur in remaining records (after the next)
                    recordIdx++;
                    for (int recordIdx2 = recordIdx; recordIdx2 < records.Count; recordIdx2++)
                    {
                        if (Check[CheckIdx].GetRecord().isInstance(records.Get(recordIdx2)))
                            Assert.Fail("Record occurs scattered throughout record chain:\n" + records.Get(recordIdx2));
                    }
                }
            }
        } */

        /* package */
        static int FindFirstRecord(IList records, Type record, int startIndex)
        {
            for (int i = startIndex; i < records.Count; i++)
            {
                if (record.Name.Equals(records[i].GetType().Name))
                    return i;
            }
            return -1;
        }

        public void CheckRecordOrder(IList records, CheckRecord[] check)
        {
            int recordIdx = 0;
            for (int checkIdx = 0; checkIdx < check.Length; checkIdx++)
            {
                recordIdx = check[checkIdx].Match(records, recordIdx);
            }
        }

        /*
        void CheckRecordOrder(List records, CheckRecord[] Check)
        {
            int CheckIndex = 0;
            for (int recordIndex = 0; recordIndex < records.Count; recordIndex++)
            {
                Record record = (Record) records.Get(recordIndex);
                if (Check[CheckIndex].GetRecord().isInstance(record))
                {
                    if (Check[CheckIndex].GetOccurance() == 'M')
                    {
                        // skip over duplicate records if multiples are allowed
                        while (recordIndex+1 < records.Count && Check[CheckIndex].GetRecord().isInstance(records.Get(recordIndex+1)))
                            recordIndex++;
    //                    lastGoodMatch = recordIndex;
                    }
                    else if (Check[CheckIndex].GetOccurance() == '1')
                    {
                        // Check next record to make sure there's not more than one
                        if (recordIndex != records.Count - 1)
                        {
                            if (Check[CheckIndex].GetRecord().isInstance(records.Get(recordIndex+1)))
                            {
                                Assert.Fail("More than one occurance of record found:\n" + records.Get(recordIndex).toString());
                            }
                        }
    //                    lastGoodMatch = recordIndex;
                    }
    //                else if (Check[CheckIndex].GetOccurance() == '0')
    //                {
    //
    //                }
                    CheckIndex++;
                }
                if (CheckIndex >= Check.Length)
                    return;
            }
            Assert.Fail("Could not Find required record: " + Check[CheckIndex]);
        } */


    }
}