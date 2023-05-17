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

namespace TestCases.HSSF.Model
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using System.Collections.Generic;
    using System.Collections;
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula.PTG;

    /**
     * Tests for {@link LinkTable}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestLinkTable
    {

        /**
         * The example file attached to bugzilla 45046 is a clear example of Name records being present
         * without an External Book (SupBook) record.  Excel has no trouble Reading this file.<br/>
         * TODO get OOO documentation updated to reflect this (that EXTERNALBOOK is optional).
         *
         * It's not clear what exact steps need to be taken in Excel to create such a workbook
         */
        [Test]
        public void TestLinkTableWithoutExternalBookRecord_bug45046()
        {
            HSSFWorkbook wb;

            try
            {
                wb = HSSFTestDataSamples.OpenSampleWorkbook("ex45046-21984.xls");
            }
            catch (Exception e)
            {
                if ("DEFINEDNAME is part of LinkTable".Equals(e.Message))
                {
                    throw new AssertionException("Identified bug 45046 b");
                }
                throw;
            }
            // some other sanity Checks
            Assert.AreEqual(3, wb.NumberOfSheets);
            String formula = wb.GetSheetAt(0).GetRow(4).GetCell(13).CellFormula;

            if ("ipcSummenproduktIntern($P5,N$6,$A$9,N$5)".Equals(formula))
            {
                // The reported symptom of this bugzilla is an earlier bug (already fixed)
                throw new AssertionException("Identified bug 41726");
                // This is observable in version 3.0
            }

            Assert.AreEqual("ipcSummenproduktIntern($C5,N$2,$A$9,N$1)", formula);
        }
        [Test]
        public void TestMultipleExternSheetRecords_bug45698()
        {
            HSSFWorkbook wb;

            try
            {
                wb = HSSFTestDataSamples.OpenSampleWorkbook("ex45698-22488.xls");
            }
            catch (Exception e)
            {
                if ("Extern sheet is part of LinkTable".Equals(e.Message))
                {
                    throw new AssertionException("Identified bug 45698");
                }
                throw;
            }
            // some other sanity Checks
            Assert.AreEqual(7, wb.NumberOfSheets);
        }
        [Test]
        public void TestExtraSheetRefs_bug45978()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex45978-extraLinkTableSheets.xls");
            /*
            ex45978-extraLinkTableSheets.xls is a cut-down version of attachment 22561.
            The original file produces the same error.

            This bug was caused by a combination of invalid sheet indexes in the EXTERNSHEET
            record, and eager Initialisation of the extern sheet references. Note - the workbook
            has 2 sheets, but the EXTERNSHEET record refers to sheet indexes 0, 1 and 2.

            Offset 0x3954 (14676)
            recordid = 0x17, size = 32
            [EXTERNSHEET]
               numOfRefs	 = 5
            refrec		 #0: extBook=0 firstSheet=0 lastSheet=0
            refrec		 #1: extBook=1 firstSheet=2 lastSheet=2
            refrec		 #2: extBook=2 firstSheet=1 lastSheet=1
            refrec		 #3: extBook=0 firstSheet=-1 lastSheet=-1
            refrec		 #4: extBook=0 firstSheet=1 lastSheet=1
            [/EXTERNSHEET]

            As it turns out, the formula in question doesn't even use externSheetIndex #1 - it
            uses #4, which Resolves to sheetIndex 1 -> 'Data'.

            It is not clear exactly what externSheetIndex #4 would refer to.  Excel seems to
            display such a formula as "''!$A2", but then complains of broken link errors.
            */

            ICell cell = wb.GetSheetAt(0).GetRow(1).GetCell(1);
            String cellFormula;
            try
            {
                cellFormula = cell.CellFormula;
            }
            catch (IndexOutOfRangeException e)
            {
                if (e.Message.Equals("Index: 2, Size: 2"))
                {
                    throw new AssertionException("Identified bug 45798");
                }
                throw e;
            }
            Assert.AreEqual("Data!$A2", cellFormula);
        }

        /**
         * This problem was visible in POI svn r763332
         * when Reading the workbook of attachment 23468 from bugzilla 47001
         */
        [Test]
        public void TestMissingExternSheetRecord_bug47001b()
        {

            Record[] recs = {
                SupBookRecord.CreateAddInFunctions(),
                new SSTRecord(),
        };
            List<Record> recList = new List<Record>(recs);
            WorkbookRecordList wrl = new WorkbookRecordList();

            LinkTable lt;
            try
            {
                lt = new LinkTable(recList, 0, wrl, new Dictionary<String, NameCommentRecord>());
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Expected an EXTERNSHEET record but got (NPOI.HSSF.record.SSTRecord)"))
                {
                    throw new AssertionException("Identified bug 47001b");
                }

                throw;
            }
            Assert.IsNotNull(lt);
        }
        [Test]
        public void TestNameCommentRecordBetweenNameRecords()
        {

            Record[] recs = {
            new NameRecord(),
            new NameCommentRecord("name1", "comment1"),
            new NameRecord(),
            new NameCommentRecord("name2", "comment2"),

            };
            List<Record> recList = new List<Record>(recs);
            WorkbookRecordList wrl = new WorkbookRecordList();
            Dictionary<String, NameCommentRecord> commentRecords = new Dictionary<String, NameCommentRecord>();

            LinkTable lt = new LinkTable(recList, 0, wrl, commentRecords);
            Assert.IsNotNull(lt);

            Assert.AreEqual(2, commentRecords.Count);
            Assert.IsTrue(recs[1] == commentRecords["name1"]); //== is intentionally not .Equals()!
            Assert.IsTrue(recs[3] == commentRecords["name2"]); //== is intentionally not .Equals()!

            Assert.AreEqual(2, lt.NumNames);
        }
        [Test]
        public void TestAddNameX()
        {
            WorkbookRecordList wrl = new WorkbookRecordList();
            wrl.Add(0, new BOFRecord());
            wrl.Add(1, new CountryRecord());
            wrl.Add(2, EOFRecord.instance);

            int numberOfSheets = 3;
            LinkTable tbl = new LinkTable(numberOfSheets, wrl);
            // creation of a new LinkTable insert two new records: SupBookRecord followed by ExternSheetRecord
            // assure they are in place:
            //    [BOFRecord]
            //    [CountryRecord]
            //    [SUPBOOK Internal References  nSheets= 3]
            //    [EXTERNSHEET]
            //    [EOFRecord]

            Assert.AreEqual(5, wrl.Records.Count);
            Assert.IsTrue(wrl[(2)] is SupBookRecord);
            SupBookRecord sup1 = (SupBookRecord)wrl[(2)];
            Assert.AreEqual(numberOfSheets, sup1.NumberOfSheets);
            Assert.IsTrue(wrl[(3)] is ExternSheetRecord);
            ExternSheetRecord extSheet = (ExternSheetRecord)wrl[(3)];
            Assert.AreEqual(0, extSheet.NumOfRefs);

            Assert.IsNull(tbl.GetNameXPtg("ISODD", -1));
            Assert.AreEqual(5, wrl.Records.Count); //still have five records

            NameXPtg namex1 = tbl.AddNameXPtg("ISODD");  // Adds two new rercords
            Assert.AreEqual(0, namex1.SheetRefIndex);
            Assert.AreEqual(0, namex1.NameIndex);
            Assert.AreEqual(namex1.ToString(), tbl.GetNameXPtg("ISODD", -1).ToString());

            // Can only find on the right sheet ref, if restricting
            Assert.AreEqual(namex1.ToString(), tbl.GetNameXPtg("ISODD", 0).ToString());
            Assert.IsNull(tbl.GetNameXPtg("ISODD", 1));
            Assert.IsNull(tbl.GetNameXPtg("ISODD", 2));
            // assure they are in place:
            //    [BOFRecord]
            //    [CountryRecord]
            //    [SUPBOOK Internal References  nSheets= 3]
            //    [SUPBOOK Add-In Functions nSheets= 1]
            //    [EXTERNALNAME .name    = ISODD]
            //    [EXTERNSHEET]
            //    [EOFRecord]

            Assert.AreEqual(7, wrl.Records.Count);
            Assert.IsTrue(wrl[(3)] is SupBookRecord);
            SupBookRecord sup2 = (SupBookRecord)wrl[(3)];
            Assert.IsTrue(sup2.IsAddInFunctions);
            Assert.IsTrue(wrl[(4)] is ExternalNameRecord);
            ExternalNameRecord ext1 = (ExternalNameRecord)wrl[(4)];
            Assert.AreEqual("ISODD", ext1.Text);
            Assert.IsTrue(wrl[(5)] is ExternSheetRecord);
            Assert.AreEqual(1, extSheet.NumOfRefs);

            //check that
            Assert.AreEqual(0, tbl.ResolveNameXIx(namex1.SheetRefIndex, namex1.NameIndex));
            Assert.AreEqual("ISODD", tbl.ResolveNameXText(namex1.SheetRefIndex, namex1.NameIndex, null));

            Assert.IsNull(tbl.GetNameXPtg("ISEVEN", -1));
            NameXPtg namex2 = tbl.AddNameXPtg("ISEVEN");  // Adds two new rercords
            Assert.AreEqual(0, namex2.SheetRefIndex);
            Assert.AreEqual(1, namex2.NameIndex);  // name index increased by one
            Assert.AreEqual(namex2.ToString(), tbl.GetNameXPtg("ISEVEN", -1).ToString());
            Assert.AreEqual(8, wrl.Records.Count);
            // assure they are in place:
            //    [BOFRecord]
            //    [CountryRecord]
            //    [SUPBOOK Internal References  nSheets= 3]
            //    [SUPBOOK Add-In Functions nSheets= 1]
            //    [EXTERNALNAME .name    = ISODD]
            //    [EXTERNALNAME .name    = ISEVEN]
            //    [EXTERNSHEET]
            //    [EOFRecord]
            Assert.IsTrue(wrl[(3)] is SupBookRecord);
            Assert.IsTrue(wrl[(4)] is ExternalNameRecord);
            Assert.IsTrue(wrl[(5)] is ExternalNameRecord);
            Assert.AreEqual("ISODD", ((ExternalNameRecord)wrl[(4)]).Text);
            Assert.AreEqual("ISEVEN", ((ExternalNameRecord)wrl[(5)]).Text);
            Assert.IsTrue(wrl[(6)] is ExternSheetRecord);
            Assert.IsTrue(wrl[(7)] is EOFRecord);

            Assert.AreEqual(0, tbl.ResolveNameXIx(namex2.SheetRefIndex, namex2.NameIndex));
            Assert.AreEqual("ISEVEN", tbl.ResolveNameXText(namex2.SheetRefIndex, namex2.NameIndex, null));

        }
    }


}