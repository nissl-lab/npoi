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
using System;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Text.RegularExpressions;
using NPOI.SS;

namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a defined named range in a SpreadsheetML workbook.
     * <p>
     * Defined names are descriptive text that is used to represents a cell, range of cells, formula, or constant value.
     * Use easy-to-understand names, such as Products, to refer to hard to understand ranges, such as <code>Sales!C20:C30</code>.
     * </p>
     * Example:
     * <pre><blockquote>
     *   XSSFWorkbook wb = new XSSFWorkbook();
     *   XSSFSheet sh = wb.CreateSheet("Sheet1");
     *
     *   //applies to the entire workbook
     *   XSSFName name1 = wb.CreateName();
     *   name1.SetNameName("FMLA");
     *   name1.SetRefersToFormula("Sheet1!$B$3");
     *
     *   //applies to Sheet1
     *   XSSFName name2 = wb.CreateName();
     *   name2.SetNameName("SheetLevelName");
     *   name2.SetComment("This name is scoped to Sheet1");
     *   name2.SetLocalSheetId(0);
     *   name2.SetRefersToFormula("Sheet1!$B$3");
     *
     * </blockquote></pre>
     *
     * @author Nick Burch
     * @author Yegor Kozlov
     */
    public class XSSFName : IName
    {
        //private static Regex isValidName = new Regex(
        //   "[\\p{IsAlphabetic}_\\\\]" +
        //   "[\\p{IsAlphabetic}0-9_.\\\\]*",
        //   RegexOptions.IgnoreCase);
        /**
         * A built-in defined name that specifies the workbook's print area
         */
        public static String BUILTIN_PRINT_AREA = "_xlnm.Print_Area";

        /**
         * A built-in defined name that specifies the row(s) or column(s) to repeat
         * at the top of each printed page.
         */
        public static String BUILTIN_PRINT_TITLE = "_xlnm.Print_Titles";

        /**
         * A built-in defined name that refers to a range Containing the criteria values
         * to be used in Applying an advanced filter to a range of data
         */
        public static String BUILTIN_CRITERIA = "_xlnm.Criteria:";


        /**
         * this defined name refers to the range Containing the filtered
         * output values resulting from Applying an advanced filter criteria to a source
         * range
         */
        public static String BUILTIN_EXTRACT = "_xlnm.Extract:";

        /**
         * ?an be one of the following
         * 1 this defined name refers to a range to which an advanced filter has been
         * applied. This represents the source data range, unfiltered.
         * 2 This defined name refers to a range to which an AutoFilter has been
         * applied
         */
        public static String BUILTIN_FILTER_DB = "_xlnm._FilterDatabase";

        /**
         * A built-in defined name that refers to a consolidation area
         */
        public static String BUILTIN_CONSOLIDATE_AREA = "_xlnm.Consolidate_Area";

        /**
         * A built-in defined name that specified that the range specified is from a database data source
         */
        public static String BUILTIN_DATABASE = "_xlnm.Database";

        /**
         * A built-in defined name that refers to a sheet title.
         */
        public static String BUILTIN_SHEET_TITLE = "_xlnm.Sheet_Title";

        private readonly XSSFWorkbook _workbook;
        private readonly CT_DefinedName _ctName;

        /**
         * Creates an XSSFName object - called internally by XSSFWorkbook.
         *
         * @param name - the xml bean that holds data represenring this defined name.
         * @param workbook - the workbook object associated with the name
         * @see NPOI.XSSF.usermodel.XSSFWorkbook#CreateName()
         */
        public XSSFName(CT_DefinedName name, XSSFWorkbook workbook)
        {
            _workbook = workbook;
            _ctName = name;
        }

        /**
         * Returns the underlying named range object
         */
        internal CT_DefinedName GetCTName()
        {
            return _ctName;
        }

        /**
         * Returns the name that will appear in the user interface for the defined name.
         *
         * @return text name of this defined name
         */
        public String NameName
        {
            get
            {
                return _ctName.name;
            }
            set
            {
                ValidateName(value);
                String oldName = NameName;
                int sheetIndex = SheetIndex;
                
                //Check to ensure no other names have the same case-insensitive name at the same scope
                foreach (XSSFName foundName in _workbook.GetNames(value))
                {
                    if (foundName != this && sheetIndex == foundName.SheetIndex)
                    {
                        String msg = "The " + (sheetIndex == -1 ? "workbook" : "sheet") + " already contains this name: " + value;
                        throw new ArgumentException(msg);
                    }
                }
                _ctName.name = value;
                //Need to update the name -> named ranges map
                _workbook.UpdateName(this, oldName);
            }
        }

        public String RefersToFormula
        {
            get
            {
                String result = _ctName.Value;
                if (result == null || result.Length < 1)
                {
                    return null;
                }
                return result;
            }
            set 
            {
                XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(_workbook);
                //validate through the FormulaParser
                FormulaParser.Parse(value, fpb, FormulaType.NamedRange, SheetIndex, -1);

                _ctName.Value = value;   
            }
        }
        public bool IsDeleted
        {
            get
            {
                String formulaText = RefersToFormula;
                if (formulaText == null)
                {
                    return false;
                }
                XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(_workbook);
                Ptg[] ptgs = FormulaParser.Parse(formulaText, fpb, FormulaType.NamedRange, SheetIndex, -1);
                return Ptg.DoesFormulaReferToDeletedCell(ptgs);
            }
        }

        /**
         * Returns the sheet index this name applies to.
         *
         * @return the sheet index this name applies to, -1 if this name applies to the entire workbook
         */
        public int SheetIndex
        {
            get
            {
                return _ctName.IsSetLocalSheetId() ? (int)_ctName.localSheetId : -1;
            }
            set 
            {
                int lastSheetIx = _workbook.NumberOfSheets - 1;
                if (value < -1 || value > lastSheetIx)
                {
                    throw new ArgumentException("Sheet index (" + value + ") is out of range" +
                            (lastSheetIx == -1 ? "" : (" (0.." + lastSheetIx + ")")));
                }

                if (value == -1)
                {
                    if (_ctName.IsSetLocalSheetId()) _ctName.UnsetLocalSheetId();
                }
                else
                {
                    _ctName.localSheetId = (uint)value;
                    _ctName.localSheetIdSpecified = true;
                }
            }
        }

        /**
         * Indicates that the defined name refers to a user-defined function.
         * This attribute is used when there is an Add-in or other code project associated with the file.
         *
         * @return <code>true</code> indicates the name refers to a function.
         */
        public bool Function
        {
            get
            {
                return _ctName.function;
            }
            set 
            {
                _ctName.function = value;
            }
        }

        public void SetFunction(bool value)
        {
            this.Function = value;
        }

        /**
         * Returns the function group index if the defined name refers to a function. The function
         * group defines the general category for the function. This attribute is used when there is
         * an Add-in or other code project associated with the file.
         *
         * @return the function group index that defines the general category for the function
         */
        public int FunctionGroupId
        {
            get
            {
                return (int)_ctName.functionGroupId;
            }
            set 
            {
                _ctName.functionGroupId = (uint)value;
            }
        }

        /**
         * Get the sheets name which this named range is referenced to
         *
         * @return sheet name, which this named range referred to.
         * Empty string if the referenced sheet name weas not found.
         */
        public String SheetName
        {
            get
            {
                if (_ctName.IsSetLocalSheetId())
                {
                    // Given as explicit sheet id
                    int sheetId = (int)_ctName.localSheetId;
                    return _workbook.GetSheetName(sheetId);
                }
                String ref1 = RefersToFormula;
                AreaReference areaRef = new AreaReference(ref1);
                return areaRef.FirstCell.SheetName;
            }
        }

        /**
         * Is the name refers to a user-defined function ?
         *
         * @return <code>true</code> if this name refers to a user-defined function
         */
        public bool IsFunctionName
        {
            get
            {
                return this.Function;
            }
        }

        /**
         * Returns the comment the user provided when the name was Created.
         *
         * @return the user comment for this named range
         */
        public String Comment
        {
            get
            {
                return _ctName.comment;
            }
            set 
            {
                _ctName.comment = value;
            }
        }



        public override int GetHashCode()
        {
            return _ctName.ToString().GetHashCode();
        }

        /**
         * Compares this name to the specified object.
         * The result is <code>true</code> if the argument is XSSFName and the
         * underlying CTDefinedName bean Equals to the CTDefinedName representing this name
         *
         * @param   o   the object to compare this <code>XSSFName</code> against.
         * @return  <code>true</code> if the <code>XSSFName </code>are Equal;
         *          <code>false</code> otherwise.
         */

        public override bool Equals(Object o)
        {
            if (o == this) return true;

            if (!(o is XSSFName)) return false;

            XSSFName cf = (XSSFName)o;
            return _ctName.name == cf.GetCTName().name && _ctName.localSheetId == cf.GetCTName().localSheetId && _ctName.Value==cf.RefersToFormula ;
        }

        /**
         * https://support.office.com/en-us/article/Define-and-use-names-in-formulas-4D0F13AC-53B7-422E-AFD2-ABD7FF379C64#bmsyntax_rules_for_names
         * 
         * Valid characters:
         *   First character: { letter | underscore | backslash }
         *   Remaining characters: { letter | number | period | underscore }
         *   
         * Cell shorthand: cannot be { "C" | "c" | "R" | "r" }
         * 
         * Cell references disallowed: cannot be a cell reference $A$1 or R1C1
         * 
         * Spaces are not valid (follows from valid characters above)
         * 
         * Name length: (XSSF-specific?) 255 characters maximum
         * 
         * Case sensitivity: all names are case-insensitive
         * 
         * Uniqueness: must be unique (for names with the same scope)
         *
         * @param name
         */
        private static void ValidateName(string name)
        {
            /* equivalent to:
            Pattern.compile(
                    "[\\p{IsAlphabetic}_]" +
                    "[\\p{IsAlphabetic}0-9_\\\\]*",
                    Pattern.CASE_INSENSITIVE).matcher(name).matches();
            \p{IsAlphabetic} doesn't work on Java 6, and other regex-based character classes don't work on unicode
            thus we are stuck with Character.isLetter (for now).
            */

            if (name.Length == 0)
            {
                throw new ArgumentException("Name cannot be blank");
            }
            if (name.Length > 255)
            {
                throw new ArgumentException("Invalid name: '" + name + "': cannot exceed 255 characters in length");
            }
            if (name.Equals("R", StringComparison.OrdinalIgnoreCase) || name.Equals("C", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Invalid name: '" + name + "': cannot be special shorthand R or C");
            }
            // is first character valid?
            char c = name[0];
            string allowedSymbols = "_\\";
            bool characterIsValid = (char.IsLetter(c) || allowedSymbols.IndexOf(c) != -1);
            if (!characterIsValid)
            {
                throw new ArgumentException("Invalid name: '" + name + "': first character must be underscore or a letter");
            }

            // are all other characters valid?
            allowedSymbols = "_.\\"; //backslashes needed for unicode escape
            foreach (char ch in name.ToCharArray())
            {
                characterIsValid = (char.IsLetterOrDigit(ch) || allowedSymbols.IndexOf(ch) != -1);
                if (!characterIsValid)
                {
                    throw new ArgumentException("Invalid name: '" + name + "': name must be letter, digit, period, or underscore");
                }
            }

            // Is the name a valid $A$1 cell reference
            // Because $, :, and ! are disallowed characters, A1-style references become just a letter-number combination
            if (Regex.IsMatch(name, "[A-Za-z]+\\d+"))
            {
                string col = Regex.Replace(name, "\\d", "");
                string row = Regex.Replace(name, "[A-Za-z]", "");
                if (CellReference.CellReferenceIsWithinRange(col, row, SpreadsheetVersion.EXCEL97))
                {
                    throw new ArgumentException("Invalid name: '" + name + "': cannot be $A$1-style cell reference");
                }
            }

            // Is the name a valid R1C1 cell reference?
            if (Regex.IsMatch(name, "[Rr]\\d+[Cc]\\d+"))
            {
                throw new ArgumentException("Invalid name: '" + name + "': cannot be R1C1-style cell reference");
            }
        }
    }
}
