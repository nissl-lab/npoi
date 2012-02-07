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
         * <li> this defined name refers to a range to which an advanced filter has been
         * applied. This represents the source data range, unfiltered.
         * <li> This defined name refers to a range to which an AutoFilter has been
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

        private XSSFWorkbook _workbook;
        private CT_DefinedName _ctName;

        /**
         * Creates an XSSFName object - called internally by XSSFWorkbook.
         *
         * @param name - the xml bean that holds data represenring this defined name.
         * @param workbook - the workbook object associated with the name
         * @see NPOI.XSSF.usermodel.XSSFWorkbook#CreateName()
         */
        protected XSSFName(CT_DefinedName name, XSSFWorkbook workbook)
        {
            _workbook = workbook;
            _ctName = name;
        }

        /**
         * Returns the underlying named range object
         */
        protected CT_DefinedName GetCTName()
        {
            return _ctName;
        }

        /**
         * Returns the name that will appear in the user interface for the defined name.
         *
         * @return text name of this defined name
         */
        public String GetNameName()
        {
            return _ctName.name;
        }

        /**
         * Sets the name that will appear in the user interface for the defined name.
         * Names must begin with a letter or underscore, not contain spaces and be unique across the workbook.
         *
         * <p>
         * A name must always be unique within its scope. POI prevents you from defining a name that is not unique
         * within its scope. However you can use the same name in different scopes. Example:
         * <pre><blockquote>
         * //by default names are workbook-global
         * XSSFName name;
         * name = workbook.CreateName();
         * name.SetNameName("sales_08");
         *
         * name = workbook.CreateName();
         * name.SetNameName("sales_08"); //will throw an exception: "The workbook already Contains this name (case-insensitive)"
         *
         * //create sheet-level name
         * name = workbook.CreateName();
         * name.SetSheetIndex(0); //the scope of the name is the first sheet
         * name.SetNameName("sales_08");  //ok
         *
         * name = workbook.CreateName();
         * name.SetSheetIndex(0);
         * name.SetNameName("sales_08");  //will throw an exception: "The sheet already Contains this name (case-insensitive)"
         *
         * </blockquote></pre>
        * </p>
         * @param name name of this defined name
         * @throws ArgumentException if the name is invalid or the workbook already Contains this name (case-insensitive)
         */
        public void SetNameName(String name)
        {
            validateName(name);

            int sheetIndex = GetSheetIndex();

            //Check to ensure no other names have the same case-insensitive name
            for (int i = 0; i < _workbook.GetNumberOfNames(); i++)
            {
                XSSFName nm = _workbook.GetNameAt(i);
                if (nm != this)
                {
                    if (name.Equals(nm.GetNameName(), StringComparison.InvariantCultureIgnoreCase) && sheetIndex == nm.GetSheetIndex())
                    {
                        String msg = "The " + (sheetIndex == -1 ? "workbook" : "sheet") + " already Contains this name: " + name;
                        throw new ArgumentException(msg);
                    }
                }
            }
            _ctName.name = name;
        }

        public String GetRefersToFormula()
        {
            String result = _ctName.Value;
            if (result == null || result.Length < 1)
            {
                return null;
            }
            return result;
        }

        public void SetRefersToFormula(String formulaText)
        {
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(_workbook);
            //validate through the FormulaParser
            FormulaParser.Parse(formulaText, fpb, FormulaType.NAMEDRANGE, GetSheetIndex());

            _ctName.Value = formulaText;
        }

        public bool IsDeleted()
        {
            String formulaText = GetRefersToFormula();
            if (formulaText == null)
            {
                return false;
            }
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(_workbook);
            Ptg[] ptgs = FormulaParser.Parse(formulaText, fpb, FormulaType.NAMEDRANGE, GetSheetIndex());
            return Ptg.DoesFormulaReferToDeletedCell(ptgs);
        }

        /**
         * Tell Excel that this name applies to the worksheet with the specified index instead of the entire workbook.
         *
         * @param index the sheet index this name applies to, -1 unSets this property making the name workbook-global
         */
        public void SetSheetIndex(int index)
        {
            int lastSheetIx = _workbook.GetNumberOfSheets() - 1;
            if (index < -1 || index > lastSheetIx)
            {
                throw new ArgumentException("Sheet index (" + index + ") is out of range" +
                        (lastSheetIx == -1 ? "" : (" (0.." + lastSheetIx + ")")));
            }

            if (index == -1)
            {
                if (_ctName.IsSetLocalSheetId()) _ctName.UnsetLocalSheetId();
            }
            else
            {
                _ctName.localSheetId = (uint)index;
            }
        }

        /**
         * Returns the sheet index this name applies to.
         *
         * @return the sheet index this name applies to, -1 if this name applies to the entire workbook
         */
        public int GetSheetIndex()
        {
            return _ctName.IsSetLocalSheetId() ? (int)_ctName.localSheetId : -1;
        }

        /**
         * Indicates that the defined name refers to a user-defined function.
         * This attribute is used when there is an Add-in or other code project associated with the file.
         *
         * @param value <code>true</code> indicates the name refers to a function.
         */
        public void SetFunction(bool value)
        {
            _ctName.function = value;
        }

        /**
         * Indicates that the defined name refers to a user-defined function.
         * This attribute is used when there is an Add-in or other code project associated with the file.
         *
         * @return <code>true</code> indicates the name refers to a function.
         */
        public bool GetFunction()
        {
            return _ctName.function;
        }

        /**
         * Specifies the function group index if the defined name refers to a function. The function
         * group defines the general category for the function. This attribute is used when there is
         * an Add-in or other code project associated with the file.
         *
         * @param functionGroupId the function group index that defines the general category for the function
         */
        public void SetFunctionGroupId(int functionGroupId)
        {
            _ctName.functionGroupId = (uint)functionGroupId;
        }

        /**
         * Returns the function group index if the defined name refers to a function. The function
         * group defines the general category for the function. This attribute is used when there is
         * an Add-in or other code project associated with the file.
         *
         * @return the function group index that defines the general category for the function
         */
        public int GetFunctionGroupId()
        {
            return (int)_ctName.functionGroupId;
        }

        /**
         * Get the sheets name which this named range is referenced to
         *
         * @return sheet name, which this named range referred to.
         * Empty string if the referenced sheet name weas not found.
         */
        public String GetSheetName()
        {
            if (_ctName.IsSetLocalSheetId())
            {
                // Given as explicit sheet id
                int sheetId = (int)_ctName.localSheetId;
                return _workbook.GetSheetName(sheetId);
            }
            String ref1 = GetRefersToFormula();
            AreaReference areaRef = new AreaReference(ref1);
            return areaRef.FirstCell.SheetName;
        }

        /**
         * Is the name refers to a user-defined function ?
         *
         * @return <code>true</code> if this name refers to a user-defined function
         */
        public bool IsFunctionName()
        {
            return GetFunction();
        }

        /**
         * Returns the comment the user provided when the name was Created.
         *
         * @return the user comment for this named range
         */
        public String GetComment()
        {
            return _ctName.comment;
        }

        /**
         * Specifies the comment the user provided when the name was Created.
         *
         * @param comment  the user comment for this named range
         */
        public void SetComment(String comment)
        {
            _ctName.comment = comment;
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
            return _ctName.ToString().Equals(cf.GetCTName().ToString());
        }

        private static void validateName(String name)
        {
            if (name.Length == 0) throw new ArgumentException("Name cannot be blank");
            char c = name[0];
            if (!(c == '_' || Char.IsLetter(c)) || name.IndexOf(' ') != -1)
            {
                throw new ArgumentException("Invalid name: '" + name + "'; Names must begin with a letter or underscore and not contain spaces");
            }
        }
    }
}
