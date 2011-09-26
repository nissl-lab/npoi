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

namespace NPOI.HSSF.Model
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TestCases.HSSF;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using System.Collections.Generic;
    using System.Collections;

    /**
     * Tests for {@link LinkTable}
     *
     * @author Josh Micich
     */
    [TestClass]
    public class TestLinkTable
    {

        /**
         * The example file attached to bugzilla 45046 is a clear example of Name records being present
         * without an External Book (SupBook) record.  Excel has no trouble Reading this file.<br/>
         * TODO get OOO documentation updated to reflect this (that EXTERNALBOOK is optional).
         *
         * It's not clear what exact steps need to be taken in Excel to create such a workbook
         */
        [TestMethod]
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
                    throw new AssertFailedException("Identified bug 45046 b");
                }
                throw e;
            }
            // some other sanity Checks
            Assert.AreEqual(3, wb.NumberOfSheets);
            String formula = wb.GetSheetAt(0).GetRow(4).GetCell(13).CellFormula;

            if ("ipcSummenproduktIntern($P5,N$6,$A$9,N$5)".Equals(formula))
            {
                // The reported symptom of this bugzilla is an earlier bug (already fixed)
                throw new AssertFailedException("Identified bug 41726");
                // This is observable in version 3.0
            }

            Assert.AreEqual("ipcSummenproduktIntern($C5,N$2,$A$9,N$1)", formula);
        }
        [TestMethod]
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
                    throw new AssertFailedException("Identified bug 45698");
                }
                throw e;
            }
            // some other sanity Checks
            Assert.AreEqual(7, wb.NumberOfSheets);
        }
        [TestMethod]
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
                    throw new AssertFailedException("Identified bug 45798");
                }
                throw e;
            }
            Assert.AreEqual("Data!$A2", cellFormula);
        }

        /**
         * This problem was visible in POI svn r763332
         * when Reading the workbook of attachment 23468 from bugzilla 47001
         */
        [TestMethod]
        public void TestMissingExternSheetRecord_bug47001b() {
		
		Record[] recs = {
				SupBookRecord.CreateAddInFunctions(),
				new SSTRecord(),
		};
		List<Record> recList =  new List<Record>(recs);
		WorkbookRecordList wrl = new WorkbookRecordList();
		
		LinkTable lt;
		try {
			lt = new LinkTable(recList, 0, wrl, new Dictionary<String, NameCommentRecord>());
		} catch (Exception e) {
			if (e.Message.Equals("Expected an EXTERNSHEET record but got (NPOI.HSSF.record.SSTRecord)")) {
				throw new AssertFailedException("Identified bug 47001b");
			}
		
			throw e;
		}
		Assert.IsNotNull(lt);
	}
        [TestMethod]
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
    }


}