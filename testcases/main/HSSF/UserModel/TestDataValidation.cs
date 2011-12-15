/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

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
    using System.IO;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.Util;
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;
    using NPOI.HSSF.EventModel;
    using NPOI.SS.UserModel;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TestCases.HSSF;
    using NPOI.HSSF.Util;

    /**
     * Class for Testing Excel's data validation mechanism
     *
     * @author Dragos Buleandra ( dragos.buleandra@trade2b.ro )
     */
    [TestClass]
    public class TestDataValidation
    {

        ///** Convenient access to ERROR_STYLE constants */
        ///*package*/
        //static HSSFDataValidation.ERRORSTYLE ES = HSSFDataValidation.ERRORSTYLE.INFO;
        ///** Convenient access to OPERATOR constants */
        ///*package*/
        //static DVConstraint.ValidationType VT = null;
        ///** Convenient access to OPERATOR constants */
        ///*package*/
        //static DVConstraint.OperatorType OP = null;

        private static void log(String msg)
        {
            //if (false) { // successful Tests should be silent
            //    System.out.println(msg);
            //}      
        }

        private class ValidationAdder
        {

            private NPOI.SS.UserModel.ICellStyle _style_1;
            private NPOI.SS.UserModel.ICellStyle _style_2;
            private int _validationType;
            private HSSFSheet _sheet;
            private int _currentRowIndex;
            private NPOI.SS.UserModel.ICellStyle _cellStyle;

            public ValidationAdder(NPOI.SS.UserModel.ISheet fSheet, NPOI.SS.UserModel.ICellStyle style_1, NPOI.SS.UserModel.ICellStyle style_2,
                    NPOI.SS.UserModel.ICellStyle cellStyle, int validationType)
            {
                _sheet = (HSSFSheet)fSheet;
                _style_1 = style_1;
                _style_2 = style_2;
                _cellStyle = cellStyle;
                _validationType = validationType;
                _currentRowIndex = fSheet.PhysicalNumberOfRows;
            }
            public void AddValidation(int operatorType, String firstFormula, String secondFormula,
                    HSSFDataValidation.ERRORSTYLE errorStyle, String ruleDescr, String promptDescr,
                    bool allowEmpty, bool inputBox, bool errorBox)
            {
                String[] explicitListValues = null;

                AddValidationInternal(operatorType, firstFormula, secondFormula, errorStyle, ruleDescr,
                        promptDescr, allowEmpty, inputBox, errorBox, true,
                        explicitListValues);
            }

            private void AddValidationInternal(int operatorType, String firstFormula,
                    String secondFormula, HSSFDataValidation.ERRORSTYLE errorStyle, String ruleDescr, String promptDescr,
                    bool allowEmpty, bool inputBox, bool errorBox, bool suppressDropDown,
                    String[] explicitListValues)
            {
                int rowNum = _currentRowIndex++;

                DVConstraint dc = CreateConstraint(operatorType, firstFormula, secondFormula, explicitListValues);

                HSSFDataValidation dv = new HSSFDataValidation(new CellRangeAddressList(rowNum, rowNum, 0, 0), dc);

                dv.EmptyCellAllowed = (allowEmpty);
                dv.ErrorStyle = (errorStyle);
                dv.CreateErrorBox("Invalid Input", "Something is wrong - Check condition!");
                dv.CreatePromptBox("Validated Cell", "Allowable values have been restricted");

                dv.ShowPromptBox = (inputBox);
                dv.ShowErrorBox = (errorBox);
                dv.SuppressDropDownArrow = (suppressDropDown);


                _sheet.AddValidationData(dv);
                WriteDataValidationSettings(_sheet, _style_1, _style_2, ruleDescr, allowEmpty,
                        inputBox, errorBox);
                if (_cellStyle != null)
                {
                    IRow row = _sheet.GetRow(_sheet.PhysicalNumberOfRows - 1);
                    ICell cell = row.CreateCell(0);
                    cell.CellStyle = (_cellStyle);
                }
                WriteOtherSettings(_sheet, _style_1, promptDescr);
            }
            private DVConstraint CreateConstraint(int operatorType, String firstFormula,
                    String secondFormula, String[] explicitListValues)
            {
                if (_validationType == DVConstraint.ValidationType.LIST)
                {
                    if (explicitListValues != null)
                    {
                        return DVConstraint.CreateExplicitListConstraint(explicitListValues);
                    }
                    return DVConstraint.CreateFormulaListConstraint(firstFormula);
                }
                if (_validationType == DVConstraint.ValidationType.TIME)
                {
                    return DVConstraint.CreateTimeConstraint(operatorType, firstFormula, secondFormula);
                }
                if (_validationType == DVConstraint.ValidationType.DATE)
                {
                    return DVConstraint.CreateDateConstraint(operatorType, firstFormula, secondFormula, null);
                }
                if (_validationType == DVConstraint.ValidationType.FORMULA)
                {
                    return DVConstraint.CreateCustomFormulaConstraint(firstFormula);
                }
                return DVConstraint.CreateNumericConstraint(_validationType, operatorType, firstFormula, secondFormula);
            }
            /**
             * Writes plain text values into cells in a tabular format to form comments readable from within 
             * the spreadsheet.
             */
            private static void WriteDataValidationSettings(NPOI.SS.UserModel.ISheet sheet, NPOI.SS.UserModel.ICellStyle style_1,
                    NPOI.SS.UserModel.ICellStyle style_2, String strCondition, bool allowEmpty, bool inputBox,
                    bool errorBox)
            {
                IRow row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                // condition's string
                ICell cell = row.CreateCell(1);
                cell.CellStyle = (style_1);
                SetCellValue(cell, strCondition);
                // allow empty cells
                cell = row.CreateCell(2);
                cell.CellStyle = (style_2);
                SetCellValue(cell, ((allowEmpty) ? "yes" : "no"));
                // show input box
                cell = row.CreateCell(3);
                cell.CellStyle = (style_2);
                SetCellValue(cell, ((inputBox) ? "yes" : "no"));
                // show error box
                cell = row.CreateCell(4);
                cell.CellStyle = (style_2);
                SetCellValue(cell, ((errorBox) ? "yes" : "no"));
            }
            private static void WriteOtherSettings(NPOI.SS.UserModel.ISheet sheet, NPOI.SS.UserModel.ICellStyle style,
                    String strStettings)
            {
                IRow row = sheet.GetRow(sheet.PhysicalNumberOfRows - 1);
                ICell cell = row.CreateCell(5);
                cell.CellStyle = (style);
                SetCellValue(cell, strStettings);
            }
            public void AddListValidation(String[] explicitListValues, String listFormula, String listValsDescr,
                    bool allowEmpty, bool suppressDropDown)
            {
                String promptDescr = (allowEmpty ? "empty ok" : "not empty")
                        + ", " + (suppressDropDown ? "no drop-down" : "drop-down");
                AddValidationInternal(DVConstraint.ValidationType.LIST, listFormula, null, HSSFDataValidation.ERRORSTYLE.STOP, listValsDescr, promptDescr,
                        allowEmpty, false, true, suppressDropDown, explicitListValues);
            }
        }

        /**
         * Manages the cell styles used for formatting the output spreadsheet
         */
        private class WorkbookFormatter
        {

            private HSSFWorkbook _wb;
            private NPOI.SS.UserModel.ICellStyle _style_1;
            private NPOI.SS.UserModel.ICellStyle _style_2;
            private NPOI.SS.UserModel.ICellStyle _style_3;
            private NPOI.SS.UserModel.ICellStyle _style_4;
            private NPOI.SS.UserModel.ISheet _currentSheet;

            public WorkbookFormatter(HSSFWorkbook wb)
            {
                _wb = wb;
                _style_1 = CreateStyle(wb, HorizontalAlignment.LEFT);
                _style_2 = CreateStyle(wb, HorizontalAlignment.CENTER);
                _style_3 = CreateStyle(wb, HorizontalAlignment.CENTER, HSSFColor.GREY_25_PERCENT.index, true);
                _style_4 = CreateHeaderStyle(wb);
            }

            private static NPOI.SS.UserModel.ICellStyle CreateStyle(HSSFWorkbook wb, HorizontalAlignment h_align, short color,
                    bool bold)
            {
                IFont font = wb.CreateFont();
                if (bold)
                {
                    font.Boldweight= (short)FontBoldWeight.BOLD;
                }

                NPOI.SS.UserModel.ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.SetFont(font);
                cellStyle.FillForegroundColor = (color);
                cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
                cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.CENTER;
                cellStyle.Alignment = (h_align);
                cellStyle.BorderLeft = (CellBorderType.THIN);
                cellStyle.LeftBorderColor = (HSSFColor.BLACK.index);
                cellStyle.BorderTop = (CellBorderType.THIN);
                cellStyle.TopBorderColor = (HSSFColor.BLACK.index);
                cellStyle.BorderRight = (CellBorderType.THIN);
                cellStyle.RightBorderColor = (HSSFColor.BLACK.index);
                cellStyle.BorderBottom = (CellBorderType.THIN);
                cellStyle.BottomBorderColor = (HSSFColor.BLACK.index);

                return cellStyle;
            }

            private static NPOI.SS.UserModel.ICellStyle CreateStyle(HSSFWorkbook wb, HorizontalAlignment h_align)
            {
                return CreateStyle(wb, h_align, HSSFColor.WHITE.index, false);
            }
            private static NPOI.SS.UserModel.ICellStyle CreateHeaderStyle(HSSFWorkbook wb)
            {
                IFont font = wb.CreateFont();
                font.Color = (HSSFColor.WHITE.index);
                font.Boldweight = (short)FontBoldWeight.BOLD;

                NPOI.SS.UserModel.ICellStyle cellStyle = wb.CreateCellStyle();
                cellStyle.FillForegroundColor = (HSSFColor.BLUE_GREY.index);
                cellStyle.FillPattern = (FillPatternType.SOLID_FOREGROUND);
                cellStyle.Alignment = (HorizontalAlignment.CENTER);
                cellStyle.VerticalAlignment = (NPOI.SS.UserModel.VerticalAlignment.CENTER);
                cellStyle.BorderLeft = (CellBorderType.THIN);
                cellStyle.LeftBorderColor = (HSSFColor.WHITE.index);
                cellStyle.BorderTop = (CellBorderType.THIN);
                cellStyle.TopBorderColor = (HSSFColor.WHITE.index);
                cellStyle.BorderRight = (CellBorderType.THIN);
                cellStyle.RightBorderColor = (HSSFColor.WHITE.index);
                cellStyle.BorderBottom = (CellBorderType.THIN);
                cellStyle.BottomBorderColor = (HSSFColor.WHITE.index);
                cellStyle.SetFont(font);
                return cellStyle;
            }


            public NPOI.SS.UserModel.ISheet CreateSheet(String sheetName)
            {
                _currentSheet = _wb.CreateSheet(sheetName);
                return _currentSheet;
            }
            public void CreateDVTypeRow(String strTypeDescription)
            {
                NPOI.SS.UserModel.ISheet sheet = _currentSheet;
                IRow row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                sheet.AddMergedRegion(new CellRangeAddress(sheet.PhysicalNumberOfRows - 1, sheet.PhysicalNumberOfRows - 1, 0, 5));
                ICell cell = row.CreateCell(0);
                SetCellValue(cell, strTypeDescription);
                cell.CellStyle = (_style_3);
                row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
            }

            public void CreateHeaderRow()
            {
                NPOI.SS.UserModel.ISheet sheet = _currentSheet;
                IRow row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
                row.Height = ((short)400);
                for (int i = 0; i < 6; i++)
                {
                    row.CreateCell(i).CellStyle = (_style_4);
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
                SetCellValue(cell, "Other settings");
            }

            public ValidationAdder CreateValidationAdder(NPOI.SS.UserModel.ICellStyle cellStyle, int dataValidationType)
            {
                return new ValidationAdder(_currentSheet, _style_1, _style_2, cellStyle, dataValidationType);
            }

            public void CreateDVDescriptionRow(String strTypeDescription)
            {
                NPOI.SS.UserModel.ISheet sheet = _currentSheet;
                IRow row = sheet.GetRow(sheet.PhysicalNumberOfRows - 1);
                sheet.AddMergedRegion(new CellRangeAddress(sheet.PhysicalNumberOfRows - 1, sheet.PhysicalNumberOfRows - 1, 0, 5));
                ICell cell = row.CreateCell(0);
                SetCellValue(cell, strTypeDescription);
                cell.CellStyle = (_style_3);
                row = sheet.CreateRow(sheet.PhysicalNumberOfRows);
            }
        }


        private void AddCustomValidations(WorkbookFormatter wf)
        {
            wf.CreateSheet("Custom");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, DVConstraint.ValidationType.FORMULA);
            va.AddValidation(DVConstraint.OperatorType.BETWEEN, "ISNUMBER($A2)", null, HSSFDataValidation.ERRORSTYLE.STOP, "ISNUMBER(A2)", "Error box type = STOP", true, true, true);
            va.AddValidation(DVConstraint.OperatorType.BETWEEN, "IF(SUM(A2:A3)=5,TRUE,FALSE)", null, HSSFDataValidation.ERRORSTYLE.WARNING, "IF(SUM(A2:A3)=5,TRUE,FALSE)", "Error box type = WARNING", false, false, true);
        }

        private static void AddSimpleNumericValidations(WorkbookFormatter wf)
        {
            // data validation's number types
            wf.CreateSheet("Numbers");

            // "Whole number" validation type
            wf.CreateDVTypeRow("Whole number");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, DVConstraint.ValidationType.INTEGER);
            va.AddValidation(DVConstraint.OperatorType.BETWEEN, "2", "6", HSSFDataValidation.ERRORSTYLE.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_BETWEEN, "2", "6", HSSFDataValidation.ERRORSTYLE.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(DVConstraint.OperatorType.EQUAL, "=3+2", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Equal to (3+2)", "Error box type = WARNING", false, false, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_EQUAL, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_THAN, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Greater than 3", "-", true, false, false);
            va.AddValidation(DVConstraint.OperatorType.LESS_THAN, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_OR_EQUAL, "4", null, HSSFDataValidation.ERRORSTYLE.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(DVConstraint.OperatorType.LESS_OR_EQUAL, "4", null, HSSFDataValidation.ERRORSTYLE.STOP, "Less than or equal to 4", "-", false, true, false);

            // "Decimal" validation type
            wf.CreateDVTypeRow("Decimal");
            wf.CreateHeaderRow();

            va = wf.CreateValidationAdder(null, DVConstraint.ValidationType.DECIMAL);
            va.AddValidation(DVConstraint.OperatorType.BETWEEN, "2", "6", HSSFDataValidation.ERRORSTYLE.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_BETWEEN, "2", "6", HSSFDataValidation.ERRORSTYLE.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(DVConstraint.OperatorType.EQUAL, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Equal to 3", "Error box type = WARNING", false, false, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_EQUAL, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_THAN, "=12/6", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Greater than (12/6)", "-", true, false, false);
            va.AddValidation(DVConstraint.OperatorType.LESS_THAN, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_OR_EQUAL, "4", null, HSSFDataValidation.ERRORSTYLE.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(DVConstraint.OperatorType.LESS_OR_EQUAL, "4", null, HSSFDataValidation.ERRORSTYLE.STOP, "Less than or equal to 4", "-", false, true, false);
        }

        private static void AddListValidations(WorkbookFormatter wf, HSSFWorkbook wb)
        {
            String cellStrValue
             = "a b c d e f g h i j k l m n o p r s t u v x y z w 0 1 2 3 4 "
            + "a b c d e f g h i j k l m n o p r s t u v x y z w 0 1 2 3 4 "
            + "a b c d e f g h i j k l m n o p r s t u v x y z w 0 1 2 3 4 "
            + "a b c d e f g h i j k l m n o p r s t u v x y z w 0 1 2 3 4 ";
            String dataSheetName = "list_data";
            // "List" Data Validation type
            NPOI.SS.UserModel.ISheet fSheet = wf.CreateSheet("Lists");
            NPOI.SS.UserModel.ISheet dataSheet = wb.CreateSheet(dataSheetName);


            wf.CreateDVTypeRow("Explicit lists - list items are explicitly provided");
            wf.CreateDVDescriptionRow("Disadvantage - sum of item's Length should be less than 255 characters");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, DVConstraint.ValidationType.LIST);
            String listValsDescr = "POIFS,HSSF,HWPF,HPSF";
            String[] listVals = listValsDescr.Split(new char[]{','});
            va.AddListValidation(listVals, null, listValsDescr, false, false);
            va.AddListValidation(listVals, null, listValsDescr, false, true);
            va.AddListValidation(listVals, null, listValsDescr, true, false);
            va.AddListValidation(listVals, null, listValsDescr, true, true);



            wf.CreateDVTypeRow("Reference lists - list items are taken from others cells");
            wf.CreateDVDescriptionRow("Advantage - no restriction regarding the sum of item's Length");
            wf.CreateHeaderRow();
            va = wf.CreateValidationAdder(null, DVConstraint.ValidationType.LIST);
            String strFormula = "$A$30:$A$39";
            va.AddListValidation(null, strFormula, strFormula, false, false);

            strFormula = dataSheetName + "!$A$1:$A$10";
            va.AddListValidation(null, strFormula, strFormula, false, false);
            NPOI.SS.UserModel.IName namedRange = wb.CreateName();
            namedRange.NameName=("myName");
            namedRange.RefersToFormula=(dataSheetName + "!$A$2:$A$7");
            strFormula = "myName";
            va.AddListValidation(null, strFormula, strFormula, false, false);
            strFormula = "offset(myName, 2, 1, 4, 2)"; // Note about last param '2': 
            // - Excel expects single row or single column when entered in UI, but Process this OK otherwise
            va.AddListValidation(null, strFormula, strFormula, false, false);

            // Add list data on same sheet
            for (int i = 0; i < 10; i++)
            {
                IRow currRow = fSheet.CreateRow(i + 29);
                SetCellValue(currRow.CreateCell(0), cellStrValue);
            }
            // Add list data on another sheet
            for (int i = 0; i < 10; i++)
            {
                IRow currRow = dataSheet.CreateRow(i + 0);
                SetCellValue(currRow.CreateCell(0), "Data a" + i);
                SetCellValue(currRow.CreateCell(1), "Data b" + i);
                SetCellValue(currRow.CreateCell(2), "Data c" + i);
            }
        }

        private static void AddDateTimeValidations(WorkbookFormatter wf, HSSFWorkbook wb)
        {
            wf.CreateSheet("Dates and Times");

            IDataFormat dataFormat = wb.CreateDataFormat();
            short fmtDate = dataFormat.GetFormat("m/d/yyyy");
            short fmtTime = dataFormat.GetFormat("h:mm");
            NPOI.SS.UserModel.ICellStyle cellStyle_date = wb.CreateCellStyle();
            cellStyle_date.DataFormat=(fmtDate);
            NPOI.SS.UserModel.ICellStyle cellStyle_time = wb.CreateCellStyle();
            cellStyle_time.DataFormat=(fmtTime);

            wf.CreateDVTypeRow("Date ( cells are already formated as date - m/d/yyyy)");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(cellStyle_date, DVConstraint.ValidationType.DATE);
            va.AddValidation(DVConstraint.OperatorType.BETWEEN, "2004/01/02", "2004/01/06", HSSFDataValidation.ERRORSTYLE.STOP, "Between 1/2/2004 and 1/6/2004 ", "Error box type = STOP", true, true, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_BETWEEN, "2004/01/01", "2004/01/06", HSSFDataValidation.ERRORSTYLE.INFO, "Not between 1/2/2004 and 1/6/2004 ", "Error box type = INFO", false, true, true);
            va.AddValidation(DVConstraint.OperatorType.EQUAL, "2004/03/02", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Equal to 3/2/2004", "Error box type = WARNING", false, false, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_EQUAL, "2004/03/02", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Not equal to 3/2/2004", "-", false, false, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_THAN, "=DATEVALUE(\"4-Jul-2001\")", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Greater than DATEVALUE('4-Jul-2001')", "-", true, false, false);
            va.AddValidation(DVConstraint.OperatorType.LESS_THAN, "2004/03/02", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Less than 3/2/2004", "-", true, true, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_OR_EQUAL, "2004/03/02", null, HSSFDataValidation.ERRORSTYLE.STOP, "Greater than or equal to 3/2/2004", "Error box type = STOP", true, false, true);
            va.AddValidation(DVConstraint.OperatorType.LESS_OR_EQUAL, "2004/03/04", null, HSSFDataValidation.ERRORSTYLE.STOP, "Less than or equal to 3/4/2004", "-", false, true, false);

            // "Time" validation type
            wf.CreateDVTypeRow("Time ( cells are already formated as time - h:mm)");
            wf.CreateHeaderRow();

            va = wf.CreateValidationAdder(cellStyle_time, DVConstraint.ValidationType.TIME);
            va.AddValidation(DVConstraint.OperatorType.BETWEEN, "12:00", "16:00", HSSFDataValidation.ERRORSTYLE.STOP, "Between 12:00 and 16:00 ", "Error box type = STOP", true, true, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_BETWEEN, "12:00", "16:00", HSSFDataValidation.ERRORSTYLE.INFO, "Not between 12:00 and 16:00 ", "Error box type = INFO", false, true, true);
            va.AddValidation(DVConstraint.OperatorType.EQUAL, "13:35", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Equal to 13:35", "Error box type = WARNING", false, false, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_EQUAL, "13:35", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Not equal to 13:35", "-", false, false, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_THAN, "12:00", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Greater than 12:00", "-", true, false, false);
            va.AddValidation(DVConstraint.OperatorType.LESS_THAN, "=1/2", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Less than (1/2) -> 12:00", "-", true, true, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_OR_EQUAL, "14:00", null, HSSFDataValidation.ERRORSTYLE.STOP, "Greater than or equal to 14:00", "Error box type = STOP", true, false, true);
            va.AddValidation(DVConstraint.OperatorType.LESS_OR_EQUAL, "14:00", null, HSSFDataValidation.ERRORSTYLE.STOP, "Less than or equal to 14:00", "-", false, true, false);
        }

        private static void AddTextLengthValidations(WorkbookFormatter wf)
        {
            wf.CreateSheet("Text Lengths");
            wf.CreateHeaderRow();

            ValidationAdder va = wf.CreateValidationAdder(null, DVConstraint.ValidationType.TEXT_LENGTH);
            va.AddValidation(DVConstraint.OperatorType.BETWEEN, "2", "6", HSSFDataValidation.ERRORSTYLE.STOP, "Between 2 and 6 ", "Error box type = STOP", true, true, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_BETWEEN, "2", "6", HSSFDataValidation.ERRORSTYLE.INFO, "Not between 2 and 6 ", "Error box type = INFO", false, true, true);
            va.AddValidation(DVConstraint.OperatorType.EQUAL, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Equal to 3", "Error box type = WARNING", false, false, true);
            va.AddValidation(DVConstraint.OperatorType.NOT_EQUAL, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Not equal to 3", "-", false, false, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_THAN, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Greater than 3", "-", true, false, false);
            va.AddValidation(DVConstraint.OperatorType.LESS_THAN, "3", null, HSSFDataValidation.ERRORSTYLE.WARNING, "Less than 3", "-", true, true, false);
            va.AddValidation(DVConstraint.OperatorType.GREATER_OR_EQUAL, "4", null, HSSFDataValidation.ERRORSTYLE.STOP, "Greater than or equal to 4", "Error box type = STOP", true, false, true);
            va.AddValidation(DVConstraint.OperatorType.LESS_OR_EQUAL, "4", null, HSSFDataValidation.ERRORSTYLE.STOP, "Less than or equal to 4", "-", false, true, false);
        }
        [TestMethod]
        public void TestDataValidation1() {
		log("\nTest no. 2 - Test Excel's Data validation mechanism");
		HSSFWorkbook wb = new HSSFWorkbook();
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

		log("    Create sheet for 'Text Length' Data Validation type... ");
		AddTextLengthValidations(wf);
		log("done !");

		// Custom Validation type
		log("    Create sheet for 'Custom' Data Validation type ... ");
		AddCustomValidations(wf);
		log("done !");

		MemoryStream baos = new MemoryStream(22000);
		try {
			wb.Write(baos);
			baos.Close();
		} catch (IOException) {
			throw;
		}
		byte[] generatedContent = baos.ToArray();
		bool isSame;
		if (false) {
			// TODO - Add proof spreadsheet and compare
			Stream proofStream = HSSFTestDataSamples.OpenSampleFileStream("TestDataValidation.xls");
			isSame = CompareStreams(proofStream, generatedContent);
		}
		isSame = true;
		
		if (isSame) {
			return;
		}
        string tempdir=System.Configuration.ConfigurationManager.AppSettings["java.io.tmpdir"];
		FileStream generatedFile = File.Open(tempdir+"GeneratedTestDataValidation.xls",FileMode.OpenOrCreate);
		try {

            generatedFile.Write(generatedContent,0,generatedContent.Length);
            generatedFile.Close();
		} catch (IOException) {
            throw;
		}
	
		
	
		Console.WriteLine("This Test case has Assert.Failed because the generated file differs from proof copy '" 
				); // TODO+ proofFile.GetAbsolutePath() + "'.");
		Console.WriteLine("The cause is usually a change to this Test, or some common spreadsheet generation code.  "
				+ "The developer has to decide whether the changes were wanted or unwanted.");
		Console.WriteLine("If the changes to the generated version were unwanted, "
				+ "make the fix elsewhere (do not modify this Test or the proof spreadsheet to get the Test working).");
		Console.WriteLine("If the changes were wanted, make sure to Open the newly generated file in Excel "
				+ "and verify it manually.  The new proof file should be submitted after it is verified to be correct.");
		Console.WriteLine("");
		Console.WriteLine("One other possible (but less likely) cause of a Assert.Failed Test is a problem in the "
				+ "comparison logic used here. Perhaps some extra file regions need to be ignored.");
        Console.WriteLine("The generated file has been saved to '" + Path.GetFullPath(tempdir + "GeneratedTestDataValidation.xls") + "' for manual inspection.");
	
		Assert.Fail("Generated file differs from proof copy.  See sysout comments for details on how to fix.");
		
	}

        private static bool CompareStreams(Stream isA, byte[] generatedContent)
        {

            Stream isB = new MemoryStream(generatedContent);

            // The allowable regions where the generated file can differ from the 
            // proof should be small (i.e. much less than 1K)
            int[] allowableDifferenceRegions = { 
				0x0228, 16,  // a region of the file containing the OS username
				0x506C, 8,   // See RootProperty (super fields _seconds_2 and _days_2)
		    };
            int[] diffs = StreamUtility.DiffStreams(isA, isB, allowableDifferenceRegions);
            if (diffs == null)
            {
                return true;
            }
            Console.WriteLine("Diff from proof: ");
            for (int i = 0; i < diffs.Length; i++)
            {
                Console.WriteLine("diff at offset: 0x" + NPOI.Util.StringUtil.ToHexString(diffs[i]));
            }
            return false;
        }





        /* package */
        static void SetCellValue(ICell cell, String text)
        {
            cell.SetCellValue(new HSSFRichTextString(text));

        }
        [TestMethod]
        public void TestAddToExistingSheet()
        {

            // dvEmpty.xls is a simple one sheet workbook.  With a DataValidations header record but no 
            // DataValidations.  It's important that the example has one SHEETPROTECTION record.
            // Such a workbook can be Created in Excel (2007) by Adding datavalidation for one cell
            // and then deleting the row that contains the cell.
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("dvEmpty.xls");
            int dvRow = 0;
            HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(0);
            DVConstraint dc = DVConstraint.CreateNumericConstraint(DVConstraint.ValidationType.INTEGER, DVConstraint.OperatorType.EQUAL, "42", null);
            HSSFDataValidation dv = new HSSFDataValidation(new CellRangeAddressList(dvRow, dvRow, 0, 0), dc);

            dv.EmptyCellAllowed=(false);
            dv.ErrorStyle=(HSSFDataValidation.ERRORSTYLE.STOP);
            dv.ShowPromptBox=(true);
            dv.CreateErrorBox("Xxx", "Yyy");
            dv.SuppressDropDownArrow=(true);

            sheet.AddValidationData(dv);

            MemoryStream baos = new MemoryStream();
            try
            {
                wb.Write(baos);
            }
            catch (IOException)
            {
                throw;
            }

            byte[] wbData = baos.ToArray();

            //if (false)
            //{ // TODO (Jul 2008) fix EventRecordFactory to Process unknown records, (and DV records for that matter)

            //    ERFListener erfListener = null; // new MyERFListener();
            //    EventRecordFactory erf = new EventRecordFactory(erfListener, null);
            //    try
            //    {
            //        POIFSFileSystem fs = new POIFSFileSystem(new MemoryStream(baos.ToArray()));
            //        erf.ProcessRecords(fs.CreatePOIFSDocumentReader("Workbook"));
            //    }
            //    catch (RecordFormatException)
            //    {
            //        throw;
            //    }
            //    catch (IOException)
            //    {
            //        throw;
            //    }
            //}
            // else verify record ordering by navigating the raw bytes

            byte[] dvHeaderRecStart = { (byte)0xB2, 0x01, 0x12, 0x00, };
            int dvHeaderOffset = FindIndex(wbData, dvHeaderRecStart);
            Assert.IsTrue(dvHeaderOffset > 0);
            int nextRecIndex = dvHeaderOffset + 22;
            int nextSid
                = ((wbData[nextRecIndex + 0] << 0) & 0x00FF)
                + ((wbData[nextRecIndex + 1] << 8) & 0xFF00)
                ;
            // nextSid should be for a DVRecord.  If anything comes between the DV header record 
            // and the DV records, Excel will not be able to Open the workbook without error.

            if (nextSid == 0x0867)
            {
                throw new AssertFailedException("Identified bug 45519");
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
                bool Match = true;
                for (int j = 1; j < searchPattern.Length; j++)
                {
                    if (searchPattern[j] != largeData[i + j])
                    {
                        Match = false;
                        break;
                    }
                }
                if (Match)
                {
                    return i;
                }
            }
            return -1;
        }
    }
}