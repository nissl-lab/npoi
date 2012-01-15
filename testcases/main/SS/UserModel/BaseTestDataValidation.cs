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

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.SS;
    using NPOI.SS.Util;
    using NPOI.Util;

    /**
     * Class for Testing Excel's data validation mechanism
     *
     * @author Dragos Buleandra ( dragos.buleandra@trade2b.ro )
     */
    [TestClass]
    public abstract class BaseTestDataValidation
    {
        private ITestDataProvider _testDataProvider;

        private static POILogger log = POILogFactory.GetLogger(typeof(BaseTestDataValidation));

        protected BaseTestDataValidation(ITestDataProvider TestDataProvider)
        {
            _testDataProvider = TestDataProvider;
        }

        /** Convenient access to ERROR_STYLE constants */
        protected static DataValidation.ErrorStyle ES = null;
        /** Convenient access to OPERATOR constants */
        protected static DataValidationConstraint.ValidationType VT = null;
        /** Convenient access to OPERATOR constants */
        protected static DataValidationConstraint.OperatorType OP = null;

        private static class ValidationAdder
        {

            private CellStyle _style_1;
            private CellStyle _style_2;
            private int _validationType;
            private Sheet _sheet;
            private int _currentRowIndex;
            private CellStyle _cellStyle;

            public ValidationAdder(Sheet fSheet, CellStyle style_1, CellStyle style_2,
                    CellStyle cellStyle, int validationType)
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

                DataValidationHelper dataValidationHelper = _sheet.DataValidationHelper;
                DataValidationConstraint dc = CreateConstraint(dataValidationHelper, operatorType, firstFormula, secondFormula, explicitListValues);

                DataValidation dv = dataValidationHelper.CreateValidation(dc, new CellRangeAddressList(rowNum, rowNum, 0, 0));

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
            private DataValidationConstraint CreateConstraint(DataValidationHelper dataValidationHelper, int operatorType, String firstFormula,
                    String secondFormula, String[] explicitListValues)
            {
                if (_validationType == VT.LIST)
                {
                    if (explicitListValues != null)
                    {
                        return dataValidationHelper.CreateExplicitListConstraint(explicitListValues);
                    }
                    return dataValidationHelper.CreateFormulaListConstraint(firstFormula);
                }
                if (_validationType == VT.TIME)
                {
                    return dataValidationHelper.CreateTimeConstraint(operatorType, firstFormula, secondFormula);
                }
                if (_validationType == VT.DATE)
                {
                    return dataValidationHelper.CreateDateConstraint(operatorType, firstFormula, secondFormula, null);
                }
                if (_validationType == VT.FORMULA)
                {
                    return dataValidationHelper.CreateCustomConstraint(firstFormula);
                }

                if (_validationType == VT.INTEGER)
                {
                    return dataValidationHelper.CreateintConstraint(operatorType, firstFormula, secondFormula);
                }
                if (_validationType == VT.DECIMAL)
                {
                    return dataValidationHelper.CreateDecimalConstraint(operatorType, firstFormula, secondFormula);
                }
                if (_validationType == VT.TEXT_LENGTH)
                {
                    return dataValidationHelper.CreateTextLengthConstraint(operatorType, firstFormula, secondFormula);
                }
                return null;
            }
            /**
             * Writes plain text values into cells in a tabular format to form comments Readable from within
             * the spreadsheet.
             */
            private static void WriteDataValidationSettings(Sheet sheet, CellStyle style_1,
                    CellStyle style_2, String strCondition, bool allowEmpty, bool inputBox,
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
            private static void WriteOtherSettings(Sheet sheet, CellStyle style,
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
                AddValidationInternal(VT.LIST, listFormula, null, ES.STOP, listValsDescr, promptDescr,
                        allowEmpty, false, true, suppressDropDown, explicitListValues);
            }
        }

        private static void log(String msg)
        {
            log.Log(POILogger.INFO, msg);
        }

        /**
         * Manages the cell styles used for formatting the output spreadsheet
         */
        private class WorkbookFormatter
        {

            private Workbook _wb;
            private CellStyle _style_1;
            private CellStyle _style_2;
            private CellStyle _style_3;
            private CellStyle _style_4;
            private Sheet _currentSheet;

            public WorkbookFormatter(Workbook wb)
            {
                _wb = wb;
                _style_1 = CreateStyle(wb, CellStyle.ALIGN_LEFT);
                _style_2 = CreateStyle(wb, CellStyle.ALIGN_CENTER);
                _style_3 = CreateStyle(wb, CellStyle.ALIGN_CENTER, IndexedColors.GREY_25_PERCENT.Index, true);
                _style_4 = CreateHeaderStyle(wb);
            }

            private static CellStyle CreateStyle(Workbook wb, short h_align, short color,
                    bool bold)
            {
                Font font = wb.CreateFont();
                if (bold)
                {
                    font.Boldweight = (/*setter*/Font.BOLDWEIGHT_BOLD);
                }

                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.Font = (/*setter*/font);
                cellStyle.FillForegroundColor = (/*setter*/color);
                cellStyle.FillPattern = (/*setter*/CellStyle.SOLID_FOREGROUND);
                cellStyle.VerticalAlignment = (/*setter*/CellStyle.VERTICAL_CENTER);
                cellStyle.Alignment = (/*setter*/h_align);
                cellStyle.BorderLeft = (/*setter*/CellStyle.BORDER_THIN);
                cellStyle.LeftBorderColor = (/*setter*/IndexedColors.BLACK.Index);
                cellStyle.BorderTop = (/*setter*/CellStyle.BORDER_THIN);
                cellStyle.TopBorderColor = (/*setter*/IndexedColors.BLACK.Index);
                cellStyle.BorderRight = (/*setter*/CellStyle.BORDER_THIN);
                cellStyle.RightBorderColor = (/*setter*/IndexedColors.BLACK.Index);
                cellStyle.BorderBottom = (/*setter*/CellStyle.BORDER_THIN);
                cellStyle.BottomBorderColor = (/*setter*/IndexedColors.BLACK.Index);

                return cellStyle;
            }

            private static CellStyle CreateStyle(Workbook wb, short h_align)
            {
                return CreateStyle(wb, h_align, IndexedColors.WHITE.Index, false);
            }
            private static CellStyle CreateHeaderStyle(Workbook wb)
            {
                Font font = wb.CreateFont();
                font.Color = (/*setter*/ IndexedColors.WHITE.Index);
                font.Boldweight = (/*setter*/Font.BOLDWEIGHT_BOLD);

                ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.FillForegroundColor = (/*setter*/IndexedColors.BLUE_GREY.Index);
                cellStyle.FillPattern = (/*setter*/CellStyle.SOLID_FOREGROUND);
                cellStyle.Alignment = (/*setter*/CellStyle.ALIGN_CENTER);
                cellStyle.VerticalAlignment = (/*setter*/CellStyle.VERTICAL_CENTER);
                cellStyle.BorderLeft = (/*setter*/CellStyle.BORDER_THIN);
                cellStyle.LeftBorderColor = (/*setter*/IndexedColors.WHITE.Index);
                cellStyle.BorderTop = (/*setter*/CellStyle.BORDER_THIN);
                cellStyle.TopBorderColor = (/*setter*/IndexedColors.WHITE.Index);
                cellStyle.BorderRight = (/*setter*/CellStyle.BORDER_THIN);
                cellStyle.RightBorderColor = (/*setter*/IndexedColors.WHITE.Index);
                cellStyle.BorderBottom = (/*setter*/CellStyle.BORDER_THIN);
                cellStyle.BottomBorderColor = (/*setter*/IndexedColors.WHITE.Index);
                cellStyle.Font = (/*setter*/font);
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

            public void CreateHeaderRow() {
		 ISheet sheet = _currentSheet;
		 IRow row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
			row.Height=(/*setter*/(short) 400);
			for (int i = 0; i < 6; i++) {
				row.CreateCell(i).CellStyle=(/*setter*/_style_4);
				if (i == 2 || i == 3 || i == 4) {
					sheet.ColumnWidth=(/*setter*/i, 3500);
				} else if (i == 5) {
					sheet.ColumnWidth=(/*setter*/i, 10000);
				} else {
					sheet.ColumnWidth=(/*setter*/i, 8000);
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

            public ValidationAdder CreateValidationAdder(CellStyle cellStyle, int dataValidationType)
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

            ValidationAdder va = wf.CreateValidationAdder(null, VT.FORMULA);
            va.AddValidation(OP.BETWEEN, "ISNUMBER($A2)", null, ES.STOP, "ISNUMBER(A2)", "Error box type = STOP", true, true, true);
            va.AddValidation(OP.BETWEEN, "IF(SUM(A2:A3)=5,TRUE,FALSE)", null, ES.WARNING, "IF(SUM(A2:A3)=5,TRUE,FALSE)", "Error box type = WARNING", false, false, true);
        }

        private static void AddSimpleNumericValidations(WorkbookFormatter wf)
        {
            // data validation's number types
            wf.CreateSheet("Numbers");

            // "Whole number" validation type
            wf.CreateDVTypeRow("Whole number");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, VT.INTEGER);
            va.AddValidation(OP.BETWEEN, "2", "6", ES.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OP.NOT_BETWEEN, "2", "6", ES.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OP.EQUAL, "=3+2", null, ES.WARNING, "Equal to (3+2)", "Error box type = WARNING", false, false, true);
            va.AddValidation(OP.NOT_EQUAL, "3", null, ES.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(OP.GREATER_THAN, "3", null, ES.WARNING, "Greater than 3", "-", true, false, false);
            va.AddValidation(OP.LESS_THAN, "3", null, ES.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(OP.GREATER_OR_EQUAL, "4", null, ES.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(OP.LESS_OR_EQUAL, "4", null, ES.STOP, "Less than or equal to 4", "-", false, true, false);

            // "Decimal" validation type
            wf.CreateDVTypeRow("Decimal");
            wf.CreateHeaderRow();

            va = wf.CreateValidationAdder(null, VT.DECIMAL);
            va.AddValidation(OP.BETWEEN, "2", "6", ES.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OP.NOT_BETWEEN, "2", "6", ES.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OP.EQUAL, "3", null, ES.WARNING, "Equal to 3", "Error box type = WARNING", false, false, true);
            va.AddValidation(OP.NOT_EQUAL, "3", null, ES.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(OP.GREATER_THAN, "=12/6", null, ES.WARNING, "Greater than (12/6)", "-", true, false, false);
            va.AddValidation(OP.LESS_THAN, "3", null, ES.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(OP.GREATER_OR_EQUAL, "4", null, ES.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(OP.LESS_OR_EQUAL, "4", null, ES.STOP, "Less than or equal to 4", "-", false, true, false);
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

            ValidationAdder va = wf.CreateValidationAdder(null, VT.LIST);
            String listValsDescr = "POIFS,HSSF,HWPF,HPSF";
            String[] listVals = listValsDescr.Split(",");
            va.AddListValidation(listVals, null, listValsDescr, false, false);
            va.AddListValidation(listVals, null, listValsDescr, false, true);
            va.AddListValidation(listVals, null, listValsDescr, true, false);
            va.AddListValidation(listVals, null, listValsDescr, true, true);



            wf.CreateDVTypeRow("Reference lists - list items are taken from others cells");
            wf.CreateDVDescriptionRow("Advantage - no restriction regarding the sum of item's length");
            wf.CreateHeaderRow();
            va = wf.CreateValidationAdder(null, VT.LIST);
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

            ValidationAdder va = wf.CreateValidationAdder(cellStyle_date, VT.DATE);
            va.AddValidation(OP.BETWEEN, "2004/01/02", "2004/01/06", ES.STOP, "Between 1/2/2004 and 1/6/2004 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OP.NOT_BETWEEN, "2004/01/01", "2004/01/06", ES.INFO, "Not between 1/2/2004 and 1/6/2004 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OP.EQUAL, "2004/03/02", null, ES.WARNING, "Equal to 3/2/2004", "Error box type = WARNING", false, false, true);
            va.AddValidation(OP.NOT_EQUAL, "2004/03/02", null, ES.WARNING, "Not equal to 3/2/2004", "-", false, false, false);
            va.AddValidation(OP.GREATER_THAN, "=DATEVALUE(\"4-Jul-2001\")", null, ES.WARNING, "Greater than DATEVALUE('4-Jul-2001')", "-", true, false, false);
            va.AddValidation(OP.LESS_THAN, "2004/03/02", null, ES.WARNING, "Less than 3/2/2004", "-", true, true, false);
            va.AddValidation(OP.GREATER_OR_EQUAL, "2004/03/02", null, ES.STOP, "Greater than or equal to 3/2/2004", "Error box type = STOP", true, false, true);
            va.AddValidation(OP.LESS_OR_EQUAL, "2004/03/04", null, ES.STOP, "Less than or equal to 3/4/2004", "-", false, true, false);

            // "Time" validation type
            wf.CreateDVTypeRow("Time ( cells are already formated as time - h:mm)");
            wf.CreateHeaderRow();

            va = wf.CreateValidationAdder(cellStyle_time, VT.TIME);
            va.AddValidation(OP.BETWEEN, "12:00", "16:00", ES.STOP, "Between 12:00 and 16:00 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OP.NOT_BETWEEN, "12:00", "16:00", ES.INFO, "Not between 12:00 and 16:00 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OP.EQUAL, "13:35", null, ES.WARNING, "Equal to 13:35", "Error box type = WARNING", false, false, true);
            va.AddValidation(OP.NOT_EQUAL, "13:35", null, ES.WARNING, "Not equal to 13:35", "-", false, false, false);
            va.AddValidation(OP.GREATER_THAN, "12:00", null, ES.WARNING, "Greater than 12:00", "-", true, false, false);
            va.AddValidation(OP.LESS_THAN, "=1/2", null, ES.WARNING, "Less than (1/2) -> 12:00", "-", true, true, false);
            va.AddValidation(OP.GREATER_OR_EQUAL, "14:00", null, ES.STOP, "Greater than or equal to 14:00", "Error box type = STOP", true, false, true);
            va.AddValidation(OP.LESS_OR_EQUAL, "14:00", null, ES.STOP, "Less than or equal to 14:00", "-", false, true, false);
        }

        private static void AddTextLengthValidations(WorkbookFormatter wf)
        {
            wf.CreateSheet("Text lengths");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, VT.TEXT_LENGTH);
            va.AddValidation(OP.BETWEEN, "2", "6", ES.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(OP.NOT_BETWEEN, "2", "6", ES.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(OP.EQUAL, "3", null, ES.WARNING, "Equal to 3", "Error box type = WARNING", false, false, true);
            va.AddValidation(OP.NOT_EQUAL, "3", null, ES.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(OP.GREATER_THAN, "3", null, ES.WARNING, "Greater than 3", "-", true, false, false);
            va.AddValidation(OP.LESS_THAN, "3", null, ES.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(OP.GREATER_OR_EQUAL, "4", null, ES.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(OP.LESS_OR_EQUAL, "4", null, ES.STOP, "Less than or equal to 4", "-", false, true, false);
        }

        public void TestDataValidation()
        {
            log("\nTest no. 2 - Test Excel's Data validation mechanism");
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            WorkbookFormatter wf = new WorkbookFormatter(wb);

            log("    Create sheet for Data Validation's number types ... ");
            AddSimpleNumericValidations(wf);
            log("done !");

            log("    Create sheet for 'List' Data Validation type ... ");
            AddListValidations(wf, wb);
            log("done !");

            log("    Create sheet for 'Date' and 'Time' Data Validation types ... ");
            AddDateTimeValidations(wf, wb);
            log("done !");

            log("    Create sheet for 'Text length' Data Validation type... ");
            AddTextLengthValidations(wf);
            log("done !");

            // Custom Validation type
            log("    Create sheet for 'Custom' Data Validation type ... ");
            AddCustomValidations(wf);
            log("done !");

            wb = _testDataProvider.WriteOutAndReadBack(wb);
        }



        /* package */
        static void SetCellValue(ICell cell, String text)
        {
            cell.SetCellValue(text);

        }

    }
}