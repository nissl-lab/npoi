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

namespace TestCases.SS.UserModel
{
    using System;

    using NUnit.Framework;

    using NPOI.SS;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;
    using NPOI.HSSF.Util;

    /**
     * Class for Testing Excel's data validation mechanism
     *
     * @author Dragos Buleandra ( dragos.buleandra@trade2b.ro )
     */
    public class BaseTestDataValidation
    {
        private ITestDataProvider _testDataProvider;

        private static POILogger log = POILogFactory.GetLogger(typeof(BaseTestDataValidation));
        public BaseTestDataValidation()
            : this(HSSFITestDataProvider.Instance)
        {
        }

        protected BaseTestDataValidation(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }


        private class ValidationAdder
        {

            private ICellStyle _style_1;
            private ICellStyle _style_2;
            private int _validationType;
            private ISheet _sheet;
            private int _currentRowIndex;
            private ICellStyle _cellStyle;

            public ValidationAdder(ISheet fSheet, ICellStyle style_1, ICellStyle style_2,
                    ICellStyle cellStyle, int validationType)
            {
                _sheet = fSheet;
                _style_1 = style_1;
                _style_2 = style_2;
                _cellStyle = cellStyle;
                _validationType = validationType;
                _currentRowIndex = fSheet.PhysicalNumberOfRows;
            }
            public void AddValidation(int operatorType, String firstFormula, String secondFormula,
                    int errorStyle, String ruleDescr, String promptDescr,
                    bool allowEmpty, bool inputBox, bool errorBox)
            {
                String[] explicitListValues = null;

                AddValidationInternal(operatorType, firstFormula, secondFormula, errorStyle, ruleDescr,
                        promptDescr, allowEmpty, inputBox, errorBox, true,
                        explicitListValues);
            }

            private void AddValidationInternal(int operatorType, String firstFormula,
                    String secondFormula, int errorStyle, String ruleDescr, String promptDescr,
                    bool allowEmpty, bool inputBox, bool errorBox, bool suppressDropDown,
                    String[] explicitListValues)
            {
                int rowNum = _currentRowIndex++;

                IDataValidationHelper dataValidationHelper = _sheet.GetDataValidationHelper();
                IDataValidationConstraint dc = CreateConstraint(dataValidationHelper, operatorType, firstFormula, secondFormula, explicitListValues);

                IDataValidation dv = dataValidationHelper.CreateValidation(dc, new CellRangeAddressList(rowNum, rowNum, 0, 0));

                dv.EmptyCellAllowed = (/*setter*/allowEmpty);
                dv.ErrorStyle = (/*setter*/errorStyle);
                dv.CreateErrorBox("Invalid Input", "Something is wrong - check condition!");
                dv.CreatePromptBox("Validated Cell", "Allowable values have been restricted");

                dv.ShowPromptBox = (/*setter*/inputBox);
                dv.ShowErrorBox = (/*setter*/errorBox);
                dv.SuppressDropDownArrow = (/*setter*/suppressDropDown);


                _sheet.AddValidationData(dv);
                WriteDataValidationSettings(_sheet, _style_1, _style_2, ruleDescr, allowEmpty,
                        inputBox, errorBox);
                if (_cellStyle != null)
                {
                    IRow row = _sheet.GetRow(_sheet.PhysicalNumberOfRows - 1);
                    ICell cell = row.CreateCell(0);
                    cell.CellStyle = (/*setter*/_cellStyle);
                }
                WriteOtherSettings(_sheet, _style_1, promptDescr);
            }
            private IDataValidationConstraint CreateConstraint(IDataValidationHelper dataValidationHelper, int operatorType, String firstFormula,
                    String secondFormula, String[] explicitListValues)
            {
                if (_validationType == ValidationType.LIST)
                {
                    if (explicitListValues != null)
                    {
                        return dataValidationHelper.CreateExplicitListConstraint(explicitListValues);
                    }
                    return dataValidationHelper.CreateFormulaListConstraint(firstFormula);
                }
                if (_validationType == ValidationType.TIME)
                {
                    return dataValidationHelper.CreateTimeConstraint(operatorType, firstFormula, secondFormula);
                }
                if (_validationType == ValidationType.DATE)
                {
                    return dataValidationHelper.CreateDateConstraint(operatorType, firstFormula, secondFormula, null);
                }
                if (_validationType == ValidationType.FORMULA)
                {
                    return dataValidationHelper.CreateCustomConstraint(firstFormula);
                }

                if (_validationType == ValidationType.INTEGER)
                {
                    return dataValidationHelper.CreateintConstraint(operatorType, firstFormula, secondFormula);
                }
                if (_validationType == ValidationType.DECIMAL)
                {
                    return dataValidationHelper.CreateDecimalConstraint(operatorType, firstFormula, secondFormula);
                }
                if (_validationType == ValidationType.TEXT_LENGTH)
                {
                    return dataValidationHelper.CreateTextLengthConstraint(operatorType, firstFormula, secondFormula);
                }
                return null;
            }
            /**
             * Writes plain text values into cells in a tabular format to form comments Readable from within
             * the spreadsheet.
             */
            private static void WriteDataValidationSettings(ISheet sheet, ICellStyle style_1,
                    ICellStyle style_2, String strCondition, bool allowEmpty, bool inputBox,
                    bool errorBox)
            {
                IRow row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                // condition's string
                ICell cell = row.CreateCell(1);
                cell.CellStyle = (/*setter*/style_1);
                SetCellValue(cell, strCondition);
                // allow empty cells
                cell = row.CreateCell(2);
                cell.CellStyle = (/*setter*/style_2);
                SetCellValue(cell, ((allowEmpty) ? "yes" : "no"));
                // show input box
                cell = row.CreateCell(3);
                cell.CellStyle = (/*setter*/style_2);
                SetCellValue(cell, ((inputBox) ? "yes" : "no"));
                // show error box
                cell = row.CreateCell(4);
                cell.CellStyle = (/*setter*/style_2);
                SetCellValue(cell, ((errorBox) ? "yes" : "no"));
            }
            private static void WriteOtherSettings(ISheet sheet, ICellStyle style,
                    String strStettings)
            {
                IRow row = sheet.GetRow(sheet.PhysicalNumberOfRows - 1);
                ICell cell = row.CreateCell(5);
                cell.CellStyle = (/*setter*/style);
                SetCellValue(cell, strStettings);
            }
            public void AddListValidation(String[] explicitListValues, String listFormula, String listValsDescr,
                    bool allowEmpty, bool suppressDropDown)
            {
                String promptDescr = (allowEmpty ? "empty ok" : "not empty")
                        + ", " + (suppressDropDown ? "no drop-down" : "drop-down");
                AddValidationInternal(ValidationType.LIST, listFormula, null, ERRORSTYLE.STOP, listValsDescr, promptDescr,
                        allowEmpty, false, true, suppressDropDown, explicitListValues);
            }
        }

        private static void Log(String msg)
        {
            log.Log(POILogger.INFO, msg);
        }

        /**
         * Manages the cell styles used for formatting the output spreadsheet
         */
        private class WorkbookFormatter
        {

            private IWorkbook _wb;
            private ICellStyle _style_1;
            private ICellStyle _style_2;
            private ICellStyle _style_3;
            private ICellStyle _style_4;
            private ISheet _currentSheet;

            public WorkbookFormatter(IWorkbook wb)
            {

                _wb = wb;

                _style_1 = CreateStyle(wb, HorizontalAlignment.Left);
                _style_2 = CreateStyle(wb, HorizontalAlignment.Center);
                _style_3 = CreateStyle(wb, HorizontalAlignment.Center, HSSFColor.Grey25Percent.Index, true);
                _style_4 = CreateHeaderStyle(wb);
            }

            private static ICellStyle CreateStyle(IWorkbook wb, HorizontalAlignment h_align, short color,
                    bool bold)
            {
                IFont font = wb.CreateFont();
                font.IsBold = bold;

                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.SetFont(font);
                cellStyle.FillForegroundColor = (/*setter*/color);
                cellStyle.FillPattern = (/*setter*/FillPattern.SolidForeground);
                cellStyle.VerticalAlignment = (/*setter*/VerticalAlignment.Center);
                cellStyle.Alignment = (/*setter*/h_align);
                cellStyle.BorderLeft = (/*setter*/BorderStyle.Thin);
                cellStyle.LeftBorderColor = (/*setter*/HSSFColor.Black.Index);
                cellStyle.BorderTop = (/*setter*/BorderStyle.Thin);
                cellStyle.TopBorderColor = (/*setter*/HSSFColor.Black.Index);
                cellStyle.BorderRight = (/*setter*/BorderStyle.Thin);
                cellStyle.RightBorderColor = (/*setter*/HSSFColor.Black.Index);
                cellStyle.BorderBottom = (/*setter*/BorderStyle.Thin);
                cellStyle.BottomBorderColor = (/*setter*/HSSFColor.Black.Index);

                return cellStyle;
            }

            private static ICellStyle CreateStyle(IWorkbook wb, HorizontalAlignment h_align)
            {
                return CreateStyle(wb, h_align, HSSFColor.White.Index, false);
            }
            private static ICellStyle CreateHeaderStyle(IWorkbook wb)
            {
                IFont font = wb.CreateFont();
                font.Color = (/*setter*/ HSSFColor.White.Index);
                font.IsBold = true;

                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.FillForegroundColor = (/*setter*/HSSFColor.BlueGrey.Index);
                cellStyle.FillPattern = (/*setter*/FillPattern.SolidForeground);
                cellStyle.Alignment = (/*setter*/HorizontalAlignment.Center);
                cellStyle.VerticalAlignment = (/*setter*/VerticalAlignment.Center);
                cellStyle.BorderLeft = (/*setter*/BorderStyle.Thin);
                cellStyle.LeftBorderColor = (/*setter*/HSSFColor.White.Index);
                cellStyle.BorderTop = (/*setter*/BorderStyle.Thin);
                cellStyle.TopBorderColor = (/*setter*/HSSFColor.White.Index);
                cellStyle.BorderRight = (/*setter*/BorderStyle.Thin);
                cellStyle.RightBorderColor = (/*setter*/HSSFColor.White.Index);
                cellStyle.BorderBottom = (/*setter*/BorderStyle.Thin);
                cellStyle.BottomBorderColor = (/*setter*/HSSFColor.White.Index);
                cellStyle.SetFont(font);
                return cellStyle;
            }


            public ISheet CreateSheet(String sheetName)
            {
                _currentSheet = _wb.CreateSheet(sheetName);
                return _currentSheet;
            }
            public void CreateDVTypeRow(String strTypeDescription)
            {
                ISheet sheet = _currentSheet;
                IRow row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                sheet.AddMergedRegion(new CellRangeAddress(sheet.PhysicalNumberOfRows - 1, sheet.PhysicalNumberOfRows - 1, 0, 5));
                ICell cell = row.CreateCell(0);
                SetCellValue(cell, strTypeDescription);
                cell.CellStyle = (/*setter*/_style_3);
                row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
            }

            public void CreateHeaderRow()
            {
                ISheet sheet = _currentSheet;
                IRow row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                row.Height = (/*setter*/(short)400);
                for (int i = 0; i < 6; i++)
                {
                    row.CreateCell(i).CellStyle = (/*setter*/_style_4);
                    if (i == 2 || i == 3 || i == 4)
                    {
                        sheet.SetColumnWidth(i, 3500);
                    }
                    else if (i == 5)
                    {
                        sheet.SetColumnWidth(i, 10000);
                    }
                    else
                    {
                        sheet.SetColumnWidth(i, 8000);
                    }
                }
                ICell cell = row.GetCell(0);
                SetCellValue(cell, "Data validation cells");
                cell = row.GetCell(1);
                SetCellValue(cell, "Condition");
                cell = row.GetCell(2);
                SetCellValue(cell, "Allow blank");
                cell = row.GetCell(3);
                SetCellValue(cell, "Prompt box");
                cell = row.GetCell(4);
                SetCellValue(cell, "Error box");
                cell = row.GetCell(5);
                SetCellValue(cell, "Other Settings");
            }

            public ValidationAdder CreateValidationAdder(ICellStyle cellStyle, int dataValidationType)
            {
                return new ValidationAdder(_currentSheet, _style_1, _style_2, cellStyle, dataValidationType);
            }

            public void CreateDVDescriptionRow(String strTypeDescription)
            {
                ISheet sheet = _currentSheet;
                IRow row = sheet.GetRow(sheet.PhysicalNumberOfRows - 1);
                sheet.AddMergedRegion(new CellRangeAddress(sheet.PhysicalNumberOfRows - 1, sheet.PhysicalNumberOfRows - 1, 0, 5));
                ICell cell = row.CreateCell(0);
                SetCellValue(cell, strTypeDescription);
                cell.CellStyle = (/*setter*/_style_3);
                row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
            }
        }


        private void AddCustomValidations(WorkbookFormatter wf)
        {
            wf.CreateSheet("Custom");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, ValidationType.FORMULA);
            va.AddValidation(OperatorType.BETWEEN, "ISNUMBER($A2)", null, ERRORSTYLE.STOP, "ISNUMBER(A2)", "Error box type = STOP", true, true, true);
            va.AddValidation(OperatorType.BETWEEN, "IF(SUM(A2:A3)=5,TRUE,FALSE)", null, ERRORSTYLE.WARNING, "IF(SUM(A2:A3)=5,TRUE,FALSE)", "Error box type = WARNING", false, false, true);
        }

        private static void AddSimpleNumericValidations(WorkbookFormatter wf)
        {
            // data validation's number types
            wf.CreateSheet("Numbers");

            // "Whole number" validation type
            wf.CreateDVTypeRow("Whole number");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, ValidationType.INTEGER);
            va.AddValidation(OperatorType.BETWEEN, "2", "6", ERRORSTYLE.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OperatorType.NOT_BETWEEN, "2", "6", ERRORSTYLE.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OperatorType.EQUAL, "=3+2", null, ERRORSTYLE.WARNING, "Equal to (3+2)", "Error box type = WARNING", false, false, true);
            va.AddValidation(OperatorType.NOT_EQUAL, "3", null, ERRORSTYLE.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(OperatorType.GREATER_THAN, "3", null, ERRORSTYLE.WARNING, "Greater than 3", "-", true, false, false);
            va.AddValidation(OperatorType.LESS_THAN, "3", null, ERRORSTYLE.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(OperatorType.GREATER_OR_EQUAL, "4", null, ERRORSTYLE.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(OperatorType.LESS_OR_EQUAL, "4", null, ERRORSTYLE.STOP, "Less than or equal to 4", "-", false, true, false);

            // "Decimal" validation type
            wf.CreateDVTypeRow("Decimal");
            wf.CreateHeaderRow();

            va = wf.CreateValidationAdder(null, ValidationType.DECIMAL);
            va.AddValidation(OperatorType.BETWEEN, "2", "6", ERRORSTYLE.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OperatorType.NOT_BETWEEN, "2", "6", ERRORSTYLE.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OperatorType.EQUAL, "3", null, ERRORSTYLE.WARNING, "Equal to 3", "Error box type = WARNING", false, false, true);
            va.AddValidation(OperatorType.NOT_EQUAL, "3", null, ERRORSTYLE.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(OperatorType.GREATER_THAN, "=12/6", null, ERRORSTYLE.WARNING, "Greater than (12/6)", "-", true, false, false);
            va.AddValidation(OperatorType.LESS_THAN, "3", null, ERRORSTYLE.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(OperatorType.GREATER_OR_EQUAL, "4", null, ERRORSTYLE.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(OperatorType.LESS_OR_EQUAL, "4", null, ERRORSTYLE.STOP, "Less than or equal to 4", "-", false, true, false);
        }

        private static void AddListValidations(WorkbookFormatter wf, IWorkbook wb)
        {
            String cellStrValue
                = "a b c d e f g h i j k l m n o p r s t u v x y z w 0 1 2 3 4 "
               + "a b c d e f g h i j k l m n o p r s t u v x y z w 0 1 2 3 4 "
               + "a b c d e f g h i j k l m n o p r s t u v x y z w 0 1 2 3 4 "
               + "a b c d e f g h i j k l m n o p r s t u v x y z w 0 1 2 3 4 ";
            String dataSheetName = "list_data";
            // "List" Data Validation type
            ISheet fSheet = wf.CreateSheet("Lists");
            ISheet dataSheet = wb.CreateSheet(dataSheetName);


            wf.CreateDVTypeRow("Explicit lists - list items are explicitly provided");
            wf.CreateDVDescriptionRow("Disadvantage - sum of item's length should be less than 255 characters");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, ValidationType.LIST);
            String listValsDescr = "POIFS,HSSF,HWPF,HPSF";
            String[] listVals = listValsDescr.Split(",".ToCharArray());
            va.AddListValidation(listVals, null, listValsDescr, false, false);
            va.AddListValidation(listVals, null, listValsDescr, false, true);
            va.AddListValidation(listVals, null, listValsDescr, true, false);
            va.AddListValidation(listVals, null, listValsDescr, true, true);



            wf.CreateDVTypeRow("Reference lists - list items are taken from others cells");
            wf.CreateDVDescriptionRow("Advantage - no restriction regarding the sum of item's length");
            wf.CreateHeaderRow();
            va = wf.CreateValidationAdder(null, ValidationType.LIST);
            String strFormula = "$A$30:$A$39";
            va.AddListValidation(null, strFormula, strFormula, false, false);

            strFormula = dataSheetName + "!$A$1:$A$10";
            va.AddListValidation(null, strFormula, strFormula, false, false);
            IName namedRange = wb.CreateName();
            namedRange.NameName = (/*setter*/"myName");
            namedRange.RefersToFormula = (/*setter*/dataSheetName + "!$A$2:$A$7");
            strFormula = "myName";
            va.AddListValidation(null, strFormula, strFormula, false, false);
            strFormula = "offset(myName, 2, 1, 4, 2)"; // Note about last param '2':
            // - Excel expects single row or single column when entered in UI, but process this OK otherwise
            va.AddListValidation(null, strFormula, strFormula, false, false);

            // add list data on same sheet
            for (int i = 0; i < 10; i++)
            {
                IRow currRow = fSheet.CreateRow(i + 29);
                SetCellValue(currRow.CreateCell(0), cellStrValue);
            }
            // add list data on another sheet
            for (int i = 0; i < 10; i++)
            {
                IRow currRow = dataSheet.CreateRow(i + 0);
                SetCellValue(currRow.CreateCell(0), "Data a" + i);
                SetCellValue(currRow.CreateCell(1), "Data b" + i);
                SetCellValue(currRow.CreateCell(2), "Data c" + i);
            }
        }

        private static void AddDateTimeValidations(WorkbookFormatter wf, IWorkbook wb)
        {
            wf.CreateSheet("Dates and Times");

            IDataFormat dataFormat = wb.CreateDataFormat();
            short fmtDate = dataFormat.GetFormat("m/d/yyyy");
            short fmtTime = dataFormat.GetFormat("h:mm");
            ICellStyle cellStyle_date = wb.CreateCellStyle();
            cellStyle_date.DataFormat = (/*setter*/fmtDate);
            ICellStyle cellStyle_time = wb.CreateCellStyle();
            cellStyle_time.DataFormat = (/*setter*/fmtTime);

            wf.CreateDVTypeRow("Date ( cells are already formated as date - m/d/yyyy)");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(cellStyle_date, ValidationType.DATE);
            va.AddValidation(OperatorType.BETWEEN, "2004/01/02", "2004/01/06", ERRORSTYLE.STOP, "Between 1/2/2004 and 1/6/2004 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OperatorType.NOT_BETWEEN, "2004/01/01", "2004/01/06", ERRORSTYLE.INFO, "Not between 1/2/2004 and 1/6/2004 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OperatorType.EQUAL, "2004/03/02", null, ERRORSTYLE.WARNING, "Equal to 3/2/2004", "Error box type = WARNING", false, false, true);
            va.AddValidation(OperatorType.NOT_EQUAL, "2004/03/02", null, ERRORSTYLE.WARNING, "Not equal to 3/2/2004", "-", false, false, false);
            va.AddValidation(OperatorType.GREATER_THAN, "=DATEVALUE(\"4-Jul-2001\")", null, ERRORSTYLE.WARNING, "Greater than DATEVALUE('4-Jul-2001')", "-", true, false, false);
            va.AddValidation(OperatorType.LESS_THAN, "2004/03/02", null, ERRORSTYLE.WARNING, "Less than 3/2/2004", "-", true, true, false);
            va.AddValidation(OperatorType.GREATER_OR_EQUAL, "2004/03/02", null, ERRORSTYLE.STOP, "Greater than or equal to 3/2/2004", "Error box type = STOP", true, false, true);
            va.AddValidation(OperatorType.LESS_OR_EQUAL, "2004/03/04", null, ERRORSTYLE.STOP, "Less than or equal to 3/4/2004", "-", false, true, false);

            // "Time" validation type
            wf.CreateDVTypeRow("Time ( cells are already formated as time - h:mm)");
            wf.CreateHeaderRow();

            va = wf.CreateValidationAdder(cellStyle_time, ValidationType.TIME);
            va.AddValidation(OperatorType.BETWEEN, "12:00", "16:00", ERRORSTYLE.STOP, "Between 12:00 and 16:00 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OperatorType.NOT_BETWEEN, "12:00", "16:00", ERRORSTYLE.INFO, "Not between 12:00 and 16:00 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OperatorType.EQUAL, "13:35", null, ERRORSTYLE.WARNING, "Equal to 13:35", "Error box type = WARNING", false, false, true);
            va.AddValidation(OperatorType.NOT_EQUAL, "13:35", null, ERRORSTYLE.WARNING, "Not equal to 13:35", "-", false, false, false);
            va.AddValidation(OperatorType.GREATER_THAN, "12:00", null, ERRORSTYLE.WARNING, "Greater than 12:00", "-", true, false, false);
            va.AddValidation(OperatorType.LESS_THAN, "=1/2", null, ERRORSTYLE.WARNING, "Less than (1/2) -> 12:00", "-", true, true, false);
            va.AddValidation(OperatorType.GREATER_OR_EQUAL, "14:00", null, ERRORSTYLE.STOP, "Greater than or equal to 14:00", "Error box type = STOP", true, false, true);
            va.AddValidation(OperatorType.LESS_OR_EQUAL, "14:00", null, ERRORSTYLE.STOP, "Less than or equal to 14:00", "-", false, true, false);
        }

        private static void AddTextLengthValidations(WorkbookFormatter wf)
        {
            wf.CreateSheet("Text lengths");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, ValidationType.TEXT_LENGTH);
            va.AddValidation(OperatorType.BETWEEN, "2", "6", ERRORSTYLE.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OperatorType.NOT_BETWEEN, "2", "6", ERRORSTYLE.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OperatorType.EQUAL, "3", null, ERRORSTYLE.WARNING, "Equal to 3", "Error box type = WARNING", false, false, true);
            va.AddValidation(OperatorType.NOT_EQUAL, "3", null, ERRORSTYLE.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(OperatorType.GREATER_THAN, "3", null, ERRORSTYLE.WARNING, "Greater than 3", "-", true, false, false);
            va.AddValidation(OperatorType.LESS_THAN, "3", null, ERRORSTYLE.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(OperatorType.GREATER_OR_EQUAL, "4", null, ERRORSTYLE.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(OperatorType.LESS_OR_EQUAL, "4", null, ERRORSTYLE.STOP, "Less than or equal to 4", "-", false, true, false);
        }
        [Test]
        public void TestDataValidation()
        {
            Log("\nTest no. 2 - Test Excel's Data validation mechanism");
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            WorkbookFormatter wf = new WorkbookFormatter(wb);

            Log("    Create sheet for Data Validation's number types ... ");
            AddSimpleNumericValidations(wf);
            Log("done !");

            Log("    Create sheet for 'List' Data Validation type ... ");
            AddListValidations(wf, wb);
            Log("done !");

            Log("    Create sheet for 'Date' and 'Time' Data Validation types ... ");
            AddDateTimeValidations(wf, wb);
            Log("done !");

            Log("    Create sheet for 'Text length' Data Validation type... ");
            AddTextLengthValidations(wf);
            Log("done !");

            // Custom Validation type
            Log("    Create sheet for 'Custom' Data Validation type ... ");
            AddCustomValidations(wf);
            Log("done !");

            wb = _testDataProvider.WriteOutAndReadBack(wb);
        }



        /* package */
        static void SetCellValue(ICell cell, String text)
        {
            cell.SetCellValue(text);

        }

    }
}