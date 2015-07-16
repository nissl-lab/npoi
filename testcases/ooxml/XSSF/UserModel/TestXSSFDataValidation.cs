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
using TestCases.SS.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using NUnit.Framework;
using NPOI.SS.Util;
using System;
using System.Text;
namespace NPOI.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFDataValidation : BaseTestDataValidation
    {

        public TestXSSFDataValidation()
            : base(XSSFITestDataProvider.instance)
        {

        }
        [Test]
        public void TestAddValidations()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("DataValidations-49244.xlsx");
            ISheet sheet = workbook.GetSheetAt(0);
            List<IDataValidation> dataValidations = ((XSSFSheet)sheet).GetDataValidations();

            /**
             * 		For each validation type, there are two cells with the same validation. This Tests
             * 		application of a single validation defInition to multiple cells.
             * 		
             * 		For list ( 3 validations for explicit and 3 for formula )
             * 			- one validation that allows blank. 
             * 			- one that does not allow blank.
             * 			- one that does not show the drop down arrow.
             * 		= 2
             * 
             * 		For number validations ( integer/decimal and text length ) with 8 different types of operators.
             *		= 50  
             * 
             * 		= 52 ( Total )
             */
            Assert.AreEqual(52, dataValidations.Count);

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            int[] validationTypes = new int[] { ValidationType.INTEGER, ValidationType.DECIMAL, ValidationType.TEXT_LENGTH };

            int[] SingleOperandOperatorTypes = new int[]{
                OperatorType.LESS_THAN,OperatorType.LESS_OR_EQUAL,
                OperatorType.GREATER_THAN,OperatorType.GREATER_OR_EQUAL,
                OperatorType.EQUAL,OperatorType.NOT_EQUAL
                };
            int[] doubleOperandOperatorTypes = new int[]{
                OperatorType.BETWEEN,OperatorType.NOT_BETWEEN
        };

            decimal value = (decimal)10, value2 = (decimal)20;
            double dvalue = (double)10.001, dvalue2 = (double)19.999;
            int lastRow = sheet.LastRowNum;
            int offset = lastRow + 3;

            int lastKnownNumValidations = dataValidations.Count;

            IRow row = sheet.CreateRow(offset++);
            ICell cell = row.CreateCell(0);
            IDataValidationConstraint explicitListValidation = dataValidationHelper.CreateExplicitListConstraint(new String[] { "MA", "MI", "CA" });
            CellRangeAddressList cellRangeAddressList = new CellRangeAddressList();
            cellRangeAddressList.AddCellRangeAddress(cell.RowIndex, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex);
            IDataValidation dataValidation = dataValidationHelper.CreateValidation(explicitListValidation, cellRangeAddressList);
            SetOtherValidationParameters(dataValidation);
            sheet.AddValidationData(dataValidation);
            lastKnownNumValidations++;

            row = sheet.CreateRow(offset++);
            cell = row.CreateCell(0);

            cellRangeAddressList = new CellRangeAddressList();
            cellRangeAddressList.AddCellRangeAddress(cell.RowIndex, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex);

            ICell firstCell = row.CreateCell(1); firstCell.SetCellValue("UT");
            ICell secondCell = row.CreateCell(2); secondCell.SetCellValue("MN");
            ICell thirdCell = row.CreateCell(3); thirdCell.SetCellValue("IL");

            int rowNum = row.RowNum + 1;
            String listFormula = new StringBuilder("$B$").Append(rowNum).Append(":").Append("$D$").Append(rowNum).ToString();
            IDataValidationConstraint formulaListValidation = dataValidationHelper.CreateFormulaListConstraint(listFormula);

            dataValidation = dataValidationHelper.CreateValidation(formulaListValidation, cellRangeAddressList);
            SetOtherValidationParameters(dataValidation);
            sheet.AddValidationData(dataValidation);
            lastKnownNumValidations++;

            offset++;
            offset++;

            for (int i = 0; i < validationTypes.Length; i++)
            {
                int validationType = validationTypes[i];
                offset = offset + 2;
                IRow row0 = sheet.CreateRow(offset++);
                ICell cell_10 = row0.CreateCell(0);
                cell_10.SetCellValue(validationType == ValidationType.DECIMAL ? "Decimal " : validationType == ValidationType.INTEGER ? "int" : "Text Length");
                offset++;
                for (int j = 0; j < SingleOperandOperatorTypes.Length; j++)
                {
                    int operatorType = SingleOperandOperatorTypes[j];
                    IRow row1 = sheet.CreateRow(offset++);

                    //For int (> and >=) we add 1 extra cell for validations whose formulae reference other cells.
                    IRow row2 = i == 0 && j < 2 ? sheet.CreateRow(offset++) : null;

                    cell_10 = row1.CreateCell(0);
                    cell_10.SetCellValue(XSSFDataValidation.operatorTypeMappings[operatorType].ToString());
                    ICell cell_11 = row1.CreateCell(1);
                    ICell cell_21 = row1.CreateCell(2);
                    ICell cell_22 = i == 0 && j < 2 ? row2.CreateCell(2) : null;

                    ICell cell_13 = row1.CreateCell(3);


                    cell_13.SetCellType(CellType.Numeric);
                    cell_13.SetCellValue(validationType == ValidationType.DECIMAL ? dvalue : (double)value);


                    //First create value based validation;
                    IDataValidationConstraint constraint = dataValidationHelper.CreateNumericConstraint(validationType, operatorType, value.ToString(), null);
                    cellRangeAddressList = new CellRangeAddressList();
                    cellRangeAddressList.AddCellRangeAddress(new CellRangeAddress(cell_11.RowIndex, cell_11.RowIndex, cell_11.ColumnIndex, cell_11.ColumnIndex));
                    IDataValidation validation = dataValidationHelper.CreateValidation(constraint, cellRangeAddressList);
                    SetOtherValidationParameters(validation);
                    sheet.AddValidationData(validation);
                    Assert.AreEqual(++lastKnownNumValidations, ((XSSFSheet)sheet).GetDataValidations().Count);

                    //Now create real formula based validation.
                    String formula1 = new CellReference(cell_13.RowIndex, cell_13.ColumnIndex).FormatAsString();
                    constraint = dataValidationHelper.CreateNumericConstraint(validationType, operatorType, formula1, null);
                    if (i == 0 && j == 0)
                    {
                        cellRangeAddressList = new CellRangeAddressList();
                        cellRangeAddressList.AddCellRangeAddress(new CellRangeAddress(cell_21.RowIndex, cell_21.RowIndex, cell_21.ColumnIndex, cell_21.ColumnIndex));
                        validation = dataValidationHelper.CreateValidation(constraint, cellRangeAddressList);
                        SetOtherValidationParameters(validation);
                        sheet.AddValidationData(validation);
                        Assert.AreEqual(++lastKnownNumValidations, ((XSSFSheet)sheet).GetDataValidations().Count);

                        cellRangeAddressList = new CellRangeAddressList();
                        cellRangeAddressList.AddCellRangeAddress(new CellRangeAddress(cell_22.RowIndex, cell_22.RowIndex, cell_22.ColumnIndex, cell_22.ColumnIndex));
                        validation = dataValidationHelper.CreateValidation(constraint, cellRangeAddressList);
                        SetOtherValidationParameters(validation);
                        sheet.AddValidationData(validation);
                        Assert.AreEqual(++lastKnownNumValidations, ((XSSFSheet)sheet).GetDataValidations().Count);
                    }
                    else if (i == 0 && j == 1)
                    {
                        cellRangeAddressList = new CellRangeAddressList();
                        cellRangeAddressList.AddCellRangeAddress(new CellRangeAddress(cell_21.RowIndex, cell_21.RowIndex, cell_21.ColumnIndex, cell_21.ColumnIndex));
                        cellRangeAddressList.AddCellRangeAddress(new CellRangeAddress(cell_22.RowIndex, cell_22.RowIndex, cell_22.ColumnIndex, cell_22.ColumnIndex));
                        validation = dataValidationHelper.CreateValidation(constraint, cellRangeAddressList);
                        SetOtherValidationParameters(validation);
                        sheet.AddValidationData(validation);
                        Assert.AreEqual(++lastKnownNumValidations, ((XSSFSheet)sheet).GetDataValidations().Count);
                    }
                    else
                    {
                        cellRangeAddressList = new CellRangeAddressList();
                        cellRangeAddressList.AddCellRangeAddress(new CellRangeAddress(cell_21.RowIndex, cell_21.RowIndex, cell_21.ColumnIndex, cell_21.ColumnIndex));
                        validation = dataValidationHelper.CreateValidation(constraint, cellRangeAddressList);
                        SetOtherValidationParameters(validation);
                        sheet.AddValidationData(validation);
                        Assert.AreEqual(++lastKnownNumValidations, ((XSSFSheet)sheet).GetDataValidations().Count);
                    }
                }

                for (int j = 0; j < doubleOperandOperatorTypes.Length; j++)
                {
                    int operatorType = doubleOperandOperatorTypes[j];
                    IRow row1 = sheet.CreateRow(offset++);

                    cell_10 = row1.CreateCell(0);
                    cell_10.SetCellValue(XSSFDataValidation.operatorTypeMappings[operatorType].ToString());

                    ICell cell_11 = row1.CreateCell(1);
                    ICell cell_21 = row1.CreateCell(2);

                    ICell cell_13 = row1.CreateCell(3);
                    ICell cell_14 = row1.CreateCell(4);


                    String value1String = validationType == ValidationType.DECIMAL ? dvalue.ToString() : value.ToString();
                    cell_13.SetCellType(CellType.Numeric);
                    cell_13.SetCellValue(validationType == ValidationType.DECIMAL ? dvalue : (int)value);

                    String value2String = validationType == ValidationType.DECIMAL ? dvalue2.ToString() : value2.ToString();
                    cell_14.SetCellType(CellType.Numeric);
                    cell_14.SetCellValue(validationType == ValidationType.DECIMAL ? dvalue2 : (int)value2);


                    //First create value based validation;
                    IDataValidationConstraint constraint = dataValidationHelper.CreateNumericConstraint(validationType, operatorType, value1String, value2String);
                    cellRangeAddressList = new CellRangeAddressList();
                    cellRangeAddressList.AddCellRangeAddress(new CellRangeAddress(cell_11.RowIndex, cell_11.RowIndex, cell_11.ColumnIndex, cell_11.ColumnIndex));
                    IDataValidation validation = dataValidationHelper.CreateValidation(constraint, cellRangeAddressList);
                    SetOtherValidationParameters(validation);
                    sheet.AddValidationData(validation);
                    Assert.AreEqual(++lastKnownNumValidations, ((XSSFSheet)sheet).GetDataValidations().Count);


                    //Now create real formula based validation.
                    String formula1 = new CellReference(cell_13.RowIndex, cell_13.ColumnIndex).FormatAsString();
                    String formula2 = new CellReference(cell_14.RowIndex, cell_14.ColumnIndex).FormatAsString();
                    constraint = dataValidationHelper.CreateNumericConstraint(validationType, operatorType, formula1, formula2);
                    cellRangeAddressList = new CellRangeAddressList();
                    cellRangeAddressList.AddCellRangeAddress(new CellRangeAddress(cell_21.RowIndex, cell_21.RowIndex, cell_21.ColumnIndex, cell_21.ColumnIndex));
                    validation = dataValidationHelper.CreateValidation(constraint, cellRangeAddressList);

                    SetOtherValidationParameters(validation);
                    sheet.AddValidationData(validation);
                    Assert.AreEqual(++lastKnownNumValidations, ((XSSFSheet)sheet).GetDataValidations().Count);
                }
            }

            workbook = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            ISheet sheetAt = workbook.GetSheetAt(0);
            Assert.AreEqual(lastKnownNumValidations, ((XSSFSheet)sheetAt).GetDataValidations().Count);
        }

        protected void SetOtherValidationParameters(IDataValidation validation)
        {
            bool yesNo = true;
            validation.EmptyCellAllowed = yesNo;
            validation.ShowErrorBox = yesNo;
            validation.ShowPromptBox = yesNo;
            validation.CreateErrorBox("Error Message Title", "Error Message");
            validation.CreatePromptBox("Prompt", "Enter some value");
            validation.SuppressDropDownArrow = yesNo;
        }
        [Test]
        public void Test53965()
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
                List<IDataValidation> lst = sheet.GetDataValidations();    //<-- works
                Assert.AreEqual(0, lst.Count);

                //create the cell that will have the validation applied
                sheet.CreateRow(0).CreateCell(0);

                IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
                IDataValidationConstraint constraint = dataValidationHelper.CreateCustomConstraint("SUM($A$1:$A$1) <= 3500");
                CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
                IDataValidation validation = dataValidationHelper.CreateValidation(constraint, addressList);
                sheet.AddValidationData(validation);

                // this line caused XmlValueOutOfRangeException , see Bugzilla 3965
                lst = sheet.GetDataValidations();
                Assert.AreEqual(1, lst.Count);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestDefaultAllowBlank()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

                XSSFDataValidation validation = CreateValidation(sheet);
                sheet.AddValidationData(validation);

                List<IDataValidation> dataValidations = sheet.GetDataValidations();
                Assert.AreEqual(true, (dataValidations[0] as XSSFDataValidation).GetCTDataValidation().allowBlank);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestSetAllowBlankToFalse()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

                XSSFDataValidation validation = CreateValidation(sheet);
                validation.GetCTDataValidation().allowBlank = (/*setter*/false);

                sheet.AddValidationData(validation);

                List<IDataValidation> dataValidations = sheet.GetDataValidations();
                Assert.AreEqual(false, (dataValidations[0] as XSSFDataValidation).GetCTDataValidation().allowBlank);
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void TestSetAllowBlankToTrue()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;

                XSSFDataValidation validation = CreateValidation(sheet);
                validation.GetCTDataValidation().allowBlank = (/*setter*/true);

                sheet.AddValidationData(validation);

                List<IDataValidation> dataValidations = sheet.GetDataValidations();
                Assert.AreEqual(true, (dataValidations[0] as XSSFDataValidation).GetCTDataValidation().allowBlank);
            }
            finally
            {
                wb.Close();
            }
        }

        private XSSFDataValidation CreateValidation(XSSFSheet sheet)
        {
            //create the cell that will have the validation applied
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0);

            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();

            IDataValidationConstraint constraint = dataValidationHelper.CreateCustomConstraint("true");
            XSSFDataValidation validation = (XSSFDataValidation)dataValidationHelper.CreateValidation(constraint, new CellRangeAddressList(0, 0, 0, 0));
            return validation;
        }

    }
}


