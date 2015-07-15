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
    using System.Collections.Generic;
    using System.IO;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NUnit.Framework;
    using TestCases.HSSF;
    using TestCases.SS.UserModel;

    /**
     * Class for Testing Excel's data validation mechanism
     *
     * @author Dragos Buleandra ( dragos.buleandra@trade2b.ro )
     */
    [TestFixture]
    public class TestDataValidation : BaseTestDataValidation
    {

        public TestDataValidation()
            : base(HSSFITestDataProvider.Instance)
        {
        }


        public void AssertDataValidation(IWorkbook wb)
        {

            MemoryStream baos = new MemoryStream(22000);
            try
            {
                wb.Write(baos);
                baos.Close();
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            byte[] generatedContent = baos.ToArray();
#if !HIDE_UNREACHABLE_CODE
            bool isSame;
            if (false)
            {
                // TODO - add proof spreadsheet and compare
                Stream proofStream = HSSFTestDataSamples.OpenSampleFileStream("TestDataValidation.xls");
                isSame = CompareStreams(proofStream, generatedContent);
            }
            isSame = true;

            if (isSame)
            {
                return;
            }

            //File tempDir = new File(System.GetProperty("java.io.tmpdir"));
            string tempDir = Path.GetTempFileName();
            //File generatedFile = new File(tempDir, "GeneratedTestDataValidation.xls");
            try
            {
                FileStream fileOut = new FileStream(tempDir, FileMode.Create);
                fileOut.Write(generatedContent, 0, generatedContent.Length);
                fileOut.Close();
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }


            Console.WriteLine("This Test case has failed because the generated file differs from proof copy '"
                    ); // TODO+ proofFile.AbsolutePath + "'.");
            Console.WriteLine("The cause is usually a change to this Test, or some common spreadsheet generation code.  "
                    + "The developer has to decide whether the Changes were wanted or unwanted.");
            Console.WriteLine("If the Changes to the generated version were unwanted, "
                    + "make the fix elsewhere (do not modify this Test or the proof spreadsheet to Get the Test working).");
            Console.WriteLine("If the Changes were wanted, make sure to open the newly generated file in Excel "
                    + "and verify it manually.  The new proof file should be submitted After it is verified to be correct.");
            Console.WriteLine("");
            Console.WriteLine("One other possible (but less likely) cause of a failed Test is a problem in the "
                    + "comparison logic used here. Perhaps some extra file regions need to be ignored.");
            Console.WriteLine("The generated file has been saved to '" + tempDir + "' for manual inspection.");

            Assert.Fail("Generated file differs from proof copy.  See sysout comments for details on how to fix.");
#endif
        }

        private static bool CompareStreams(Stream isA, byte[] generatedContent)
        {

            Stream isB = new MemoryStream(generatedContent);

            // The allowable regions where the generated file can differ from the 
            // proof should be small (i.e. much less than 1K)
            int[] allowableDifferenceRegions = { 
                0x0228, 16,  // a region of the file Containing the OS username
                0x506C, 8,   // See RootProperty (super fields _seconds_2 and _days_2)
        };
            int[] diffs = StreamUtility.DiffStreams(isA, isB, allowableDifferenceRegions);
            if (diffs == null)
            {
                return true;
            }
            System.Console.Error.WriteLine("Diff from proof: ");
            for (int i = 0; i < diffs.Length; i++)
            {
                System.Console.Error.WriteLine("diff at offset: 0x" + (diffs[i]).ToString("X2"));
            }
            return false;
        }





        /* package */
        static void SetCellValue(HSSFCell cell, String text)
        {
            cell.SetCellValue(new HSSFRichTextString(text));

        }
        [Test]
        public void TestAddToExistingSheet()
        {

            // dvEmpty.xls is a simple one sheet workbook.  With a DataValidations header record but no 
            // DataValidations.  It's important that the example has one SHEETPROTECTION record.
            // Such a workbook can be Created in Excel (2007) by Adding datavalidation for one cell
            // and then deleting the row that Contains the cell.
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("dvEmpty.xls");
            int dvRow = 0;
            ISheet sheet = wb.GetSheetAt(0);
            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint dc = dataValidationHelper.CreateintConstraint(OperatorType.EQUAL, "42", null);
            IDataValidation dv = dataValidationHelper.CreateValidation(dc, new CellRangeAddressList(dvRow, dvRow, 0, 0));

            dv.EmptyCellAllowed = (/*setter*/false);
            dv.ErrorStyle = (/*setter*/ERRORSTYLE.STOP);
            dv.ShowPromptBox = (/*setter*/true);
            dv.CreateErrorBox("Xxx", "Yyy");
            dv.SuppressDropDownArrow = (/*setter*/true);

            sheet.AddValidationData(dv);

            MemoryStream baos = new MemoryStream();
            try
            {
                wb.Write(baos);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }

            byte[] wbData = baos.ToArray();

#if !HIDE_UNREACHABLE_CODE
            if (false)
            { // TODO (Jul 2008) fix EventRecordFactory to process unknown records, (and DV records for that matter)

                ERFListener erfListener = null; // new MyERFListener();
                EventRecordFactory erf = new EventRecordFactory(erfListener, null);
                try
                {
                    POIFSFileSystem fs = new POIFSFileSystem(new MemoryStream(baos.ToArray()));
                    throw new NotImplementedException("The method CreateDocumentInputStream of POIFSFileSystem is not implemented.");
                    //erf.ProcessRecords(fs.CreateDocumentInputStream("Workbook"));
                }
                catch (RecordFormatException e)
                {
                    throw new RuntimeException(e);
                }
                catch (IOException e)
                {
                    throw new RuntimeException(e);
                }
            }
            // else verify record ordering by navigating the raw bytes
#endif

            byte[] dvHeaderRecStart = { (byte)0xB2, 0x01, 0x12, 0x00, };
            int dvHeaderOffset = FindIndex(wbData, dvHeaderRecStart);
            Assert.IsTrue(dvHeaderOffset > 0);
            int nextRecIndex = dvHeaderOffset + 22;
            int nextSid
                = ((wbData[nextRecIndex + 0] << 0) & 0x00FF)
                + ((wbData[nextRecIndex + 1] << 8) & 0xFF00)
                ;
            // nextSid should be for a DVRecord.  If anything comes between the DV header record 
            // and the DV records, Excel will not be able to open the workbook without error.

            if (nextSid == 0x0867)
            {
                throw new AssertionException("Identified bug 45519");
            }
            Assert.AreEqual(DVRecord.sid, nextSid);
        }
        private int FindIndex(byte[] largeData, byte[] searchPattern)
        {
            byte firstByte = searchPattern[0];
            for (int i = 0; i < largeData.Length; i++)
            {
                if (largeData[i] != firstByte)
                {
                    continue;
                }
                bool match = true;
                for (int j = 1; j < searchPattern.Length; j++)
                {
                    if (searchPattern[j] != largeData[i + j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match)
                {
                    return i;
                }
            }
            return -1;
        }

        [Test]
        public void TestGetDataValidationsAny()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            List<IDataValidation> list = sheet.GetDataValidations();
            Assert.AreEqual(0, list.Count);

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint constraint = dataValidationHelper.CreateNumericConstraint(ValidationType.ANY,
                    OperatorType.IGNORED, null, null);
            CellRangeAddressList AddressList = new CellRangeAddressList(1, 2, 3, 4);
            IDataValidation validation = dataValidationHelper.CreateValidation(constraint, AddressList);
            validation.EmptyCellAllowed = (/*setter*/true);
            validation.CreateErrorBox("error-title", "error-text");
            validation.CreatePromptBox("prompt-title", "prompt-text");
            sheet.AddValidationData(validation);

            list = sheet.GetDataValidations(); // <-- works
            Assert.AreEqual(1, list.Count);

            HSSFDataValidation dv = list[(0)] as HSSFDataValidation;
            {
                CellRangeAddressList regions = dv.Regions;
                Assert.AreEqual(1, regions.CountRanges());

                CellRangeAddress Address = regions.GetCellRangeAddress(0);
                Assert.AreEqual(1, Address.FirstRow);
                Assert.AreEqual(2, Address.LastRow);
                Assert.AreEqual(3, Address.FirstColumn);
                Assert.AreEqual(4, Address.LastColumn);
            }
            Assert.AreEqual(true, dv.EmptyCellAllowed);
            Assert.AreEqual(false, dv.SuppressDropDownArrow);
            Assert.AreEqual(true, dv.ShowErrorBox);
            Assert.AreEqual("error-title", dv.ErrorBoxTitle);
            Assert.AreEqual("error-text", dv.ErrorBoxText);
            Assert.AreEqual(true, dv.ShowPromptBox);
            Assert.AreEqual("prompt-title", dv.PromptBoxTitle);
            Assert.AreEqual("prompt-text", dv.PromptBoxText);

            IDataValidationConstraint c = dv.ValidationConstraint;
            Assert.AreEqual(ValidationType.ANY, c.GetValidationType());
            Assert.AreEqual(OperatorType.IGNORED, c.Operator);
        }

        [Test]
        public void TestGetDataValidationsintFormula()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            List<IDataValidation> list = sheet.GetDataValidations();
            Assert.AreEqual(0, list.Count);

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint constraint = dataValidationHelper.CreateintConstraint(OperatorType.BETWEEN, "=A2",
                    "=A3");
            CellRangeAddressList AddressList = new CellRangeAddressList(0, 0, 0, 0);
            IDataValidation validation = dataValidationHelper.CreateValidation(constraint, AddressList);
            sheet.AddValidationData(validation);

            list = sheet.GetDataValidations(); // <-- works
            Assert.AreEqual(1, list.Count);

            HSSFDataValidation dv = list[(0)] as HSSFDataValidation;
            DVConstraint c = dv.Constraint;
            Assert.AreEqual(ValidationType.INTEGER, c.GetValidationType());
            Assert.AreEqual(OperatorType.BETWEEN, c.Operator);
            Assert.AreEqual("A2", c.Formula1);
            Assert.AreEqual("A3", c.Formula2);
            Assert.AreEqual(double.NaN, c.Value1);
            Assert.AreEqual(double.NaN, c.Value2);
        }

        [Test]
        public void TestGetDataValidationsintValue(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
        List<IDataValidation> list = sheet.GetDataValidations();
        Assert.AreEqual(0, list.Count);

        IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
        IDataValidationConstraint constraint = dataValidationHelper.CreateintConstraint(OperatorType.BETWEEN, "100",
                "200");
        CellRangeAddressList AddressList = new CellRangeAddressList(0, 0, 0, 0);
        IDataValidation validation = dataValidationHelper.CreateValidation(constraint, AddressList);
        sheet.AddValidationData(validation);

        list = sheet.GetDataValidations(); // <-- works
        Assert.AreEqual(1, list.Count);

        HSSFDataValidation dv = list[(0)] as HSSFDataValidation;
        DVConstraint c = dv.Constraint;
        Assert.AreEqual(ValidationType.INTEGER, c.GetValidationType());
        Assert.AreEqual(OperatorType.BETWEEN, c.Operator);
        Assert.AreEqual(null, c.Formula1);
        Assert.AreEqual(null, c.Formula2);
        Assert.AreEqual(100, c.Value1);
        Assert.AreEqual(200, c.Value2);
    }

        [Test]
        public void TestGetDataValidationsDecimal(){
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
        List<IDataValidation> list = sheet.GetDataValidations();
        Assert.AreEqual(0, list.Count);

        IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
        IDataValidationConstraint constraint = dataValidationHelper.CreateDecimalConstraint(OperatorType.BETWEEN, "=A2",
                "200");
        CellRangeAddressList AddressList = new CellRangeAddressList(0, 0, 0, 0);
        IDataValidation validation = dataValidationHelper.CreateValidation(constraint, AddressList);
        sheet.AddValidationData(validation);

        list = sheet.GetDataValidations(); // <-- works
        Assert.AreEqual(1, list.Count);

        HSSFDataValidation dv = list[(0)] as HSSFDataValidation;
        DVConstraint c = dv.Constraint;
        Assert.AreEqual(ValidationType.DECIMAL, c.GetValidationType());
        Assert.AreEqual(OperatorType.BETWEEN, c.Operator);
        Assert.AreEqual("A2", c.Formula1);
        Assert.AreEqual(null, c.Formula2);
        Assert.AreEqual(double.NaN, c.Value1);
        Assert.AreEqual(200, c.Value2);
    }

        [Test]
        public void TestGetDataValidationsDate()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            List<IDataValidation> list = sheet.GetDataValidations();
            Assert.AreEqual(0, list.Count);

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint constraint = dataValidationHelper.CreateDateConstraint(OperatorType.EQUAL,
                    "2014/10/25", null, null);
            CellRangeAddressList AddressList = new CellRangeAddressList(0, 0, 0, 0);
            IDataValidation validation = dataValidationHelper.CreateValidation(constraint, AddressList);
            sheet.AddValidationData(validation);

            list = sheet.GetDataValidations(); // <-- works
            Assert.AreEqual(1, list.Count);

            HSSFDataValidation dv = list[(0)] as HSSFDataValidation;
            DVConstraint c = dv.Constraint;
            Assert.AreEqual(ValidationType.DATE, c.GetValidationType());
            Assert.AreEqual(OperatorType.EQUAL, c.Operator);
            Assert.AreEqual(null, c.Formula1);
            Assert.AreEqual(null, c.Formula2);
            Assert.AreEqual(DateUtil.GetExcelDate(DateUtil.ParseYYYYMMDDDate("2014/10/25")), c.Value1);
            Assert.AreEqual(double.NaN, c.Value2);
        }

        [Test]
        public void TestGetDataValidationsListExplicit()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            List<IDataValidation> list = sheet.GetDataValidations();
            Assert.AreEqual(0, list.Count);

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint constraint = dataValidationHelper.CreateExplicitListConstraint(new String[] { "aaa",
                "bbb", "ccc" });
            CellRangeAddressList AddressList = new CellRangeAddressList(0, 0, 0, 0);
            IDataValidation validation = dataValidationHelper.CreateValidation(constraint, AddressList);
            validation.SuppressDropDownArrow = (/*setter*/true);
            sheet.AddValidationData(validation);

            list = sheet.GetDataValidations(); // <-- works
            Assert.AreEqual(1, list.Count);

            HSSFDataValidation dv = list[(0)] as HSSFDataValidation;
            Assert.AreEqual(true, dv.SuppressDropDownArrow);

            DVConstraint c = dv.Constraint;
            Assert.AreEqual(ValidationType.LIST, c.GetValidationType());
            Assert.AreEqual(null, c.Formula1);
            Assert.AreEqual(null, c.Formula2);
            Assert.AreEqual(double.NaN, c.Value1);
            Assert.AreEqual(double.NaN, c.Value2);
            String[] values = c.ExplicitListValues;
            Assert.AreEqual(3, values.Length);
            Assert.AreEqual("aaa", values[0]);
            Assert.AreEqual("bbb", values[1]);
            Assert.AreEqual("ccc", values[2]);
        }

        [Test]
        public void TestGetDataValidationsListFormula()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            List<IDataValidation> list = sheet.GetDataValidations();
            Assert.AreEqual(0, list.Count);

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint constraint = dataValidationHelper.CreateFormulaListConstraint("A2");
            CellRangeAddressList AddressList = new CellRangeAddressList(0, 0, 0, 0);
            IDataValidation validation = dataValidationHelper.CreateValidation(constraint, AddressList);
            validation.SuppressDropDownArrow = (/*setter*/true);
            sheet.AddValidationData(validation);

            list = sheet.GetDataValidations(); // <-- works
            Assert.AreEqual(1, list.Count);

            HSSFDataValidation dv = list[(0)] as HSSFDataValidation;
            Assert.AreEqual(true, dv.SuppressDropDownArrow);

            DVConstraint c = dv.Constraint;
            Assert.AreEqual(ValidationType.LIST, c.GetValidationType());
            Assert.AreEqual("A2", c.Formula1);
            Assert.AreEqual(null, c.Formula2);
            Assert.AreEqual(double.NaN, c.Value1);
            Assert.AreEqual(double.NaN, c.Value2);
        }

        [Test]
        public void TestGetDataValidationsFormula()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;
            List<IDataValidation> list = sheet.GetDataValidations();
            Assert.AreEqual(0, list.Count);

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint constraint = dataValidationHelper.CreateCustomConstraint("A2:A3");
            CellRangeAddressList AddressList = new CellRangeAddressList(0, 0, 0, 0);
            IDataValidation validation = dataValidationHelper.CreateValidation(constraint, AddressList);
            sheet.AddValidationData(validation);

            list = sheet.GetDataValidations(); // <-- works
            Assert.AreEqual(1, list.Count);

            HSSFDataValidation dv = list[(0)] as HSSFDataValidation;
            DVConstraint c = dv.Constraint;
            Assert.AreEqual(ValidationType.FORMULA, c.GetValidationType());
            Assert.AreEqual("A2:A3", c.Formula1);
            Assert.AreEqual(null, c.Formula2);
            Assert.AreEqual(double.NaN, c.Value1);
            Assert.AreEqual(double.NaN, c.Value2);
        }

    }

}