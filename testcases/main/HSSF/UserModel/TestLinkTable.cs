namespace TestCases.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using NUnit.Framework;

    using TestCases.HSSF;
    /**
     * Tests for LinkTable
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestLinkTable
    {

        /**
         * The example file attached to bugzilla 45046 is a Clear example of Name records being present
         * without an External Book (SupBook) record.  Excel has no trouble reading this file.<br/>
         * TODO get OOO documentation updated to reflect this (that EXTERNALBOOK is optional).
         *
         * It's not Clear what exact steps need to be taken in Excel to Create such a workbook
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
            record, and eager initialisation of the extern sheet references. Note - the worbook
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
            uses #4, which resolves to sheetIndex 1 -> 'Data'.

            It is not Clear exactly what externSheetIndex #4 would refer to.  Excel seems to
            display such a formula as "''!$A2", but then complains of broken link errors.
            */

            NPOI.SS.UserModel.ICell cell = wb.GetSheetAt(0).GetRow(1).GetCell(1);
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
    }
}