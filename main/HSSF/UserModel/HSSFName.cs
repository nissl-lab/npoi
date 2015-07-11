/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Text;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula.PTG;


    /// <summary>
    /// High Level Represantion of Named Range
    /// </summary>
    /// <remarks>@author Libin Roman (Vista Portal LDT. Developer)</remarks>
    public class HSSFName:NPOI.SS.UserModel.IName
    {
        private HSSFWorkbook book;
        private NameRecord _definedNameRec;
        private NameCommentRecord _commentRec;
        /* package */
        internal HSSFName(HSSFWorkbook book, NameRecord name):this(book, name, null)
        {
            
        }
        /// <summary>
        /// Creates new HSSFName   - called by HSSFWorkbook to Create a sheet from
        /// scratch.
        /// </summary>
        /// <param name="book">lowlevel Workbook object associated with the sheet.</param>
        /// <param name="name">the Name Record</param>
        /// <param name="comment"></param>
        internal HSSFName(HSSFWorkbook book, NameRecord name, NameCommentRecord comment)
        {
            this.book = book;
            this._definedNameRec = name;
            _commentRec = comment;
        }

        /// <summary>
        /// Gets or sets the sheets name which this named range is referenced to
        /// </summary>
        /// <value>sheet name, which this named range refered to</value>
        public String SheetName
        {
            get
            {
                String result;
                int indexToExternSheet = _definedNameRec.ExternSheetNumber;

                result = book.Workbook.FindSheetFirstNameFromExternSheet(indexToExternSheet);

                return result;
            }
            //set
            //{
            //    int sheetNumber = book.GetSheetIndex(value);

            //    int externSheetNumber = book.GetExternalSheetIndex(sheetNumber);
            //    name.ExternSheetNumber = externSheetNumber;
            //}
        }

        /// <summary>
        /// Gets or sets the name of the named range
        /// </summary>
        /// <value>named range name</value>
        public String NameName
        {
            get
            {
                String result = _definedNameRec.NameText;

                return result;
            }
            set
            {
                ValidateName(value);
                _definedNameRec.NameText = value;
                InternalWorkbook wb = book.Workbook;
                int sheetNumber = _definedNameRec.SheetNumber;

                //Check to Ensure no other names have the same case-insensitive name
                for (int i = wb.NumNames- 1; i >= 0; i--)
                {
                    NameRecord rec = wb.GetNameRecord(i);
                    if (rec != _definedNameRec)
                    {
                        if (rec.NameText.Equals(NameName, StringComparison.OrdinalIgnoreCase) && sheetNumber == rec.SheetNumber)
                        {
                            String msg = "The " + (sheetNumber == 0 ? "workbook" : "sheet") + " already contains this name: " + value;
                            _definedNameRec.NameText = (value + "(2)");
                            throw new ArgumentException(msg);
                        }
                    }
                }
                // Update our comment, if there is one
                if (_commentRec != null)
                {
                    String oldName = _commentRec.NameText;
                    _commentRec.NameText=value;
                    book.Workbook.UpdateNameCommentRecordCache(_commentRec);
                }
            }
        }
        private void ValidateName(String name)
        {
            if (name.Length == 0) throw new ArgumentException("Name cannot be blank");

            char c = name[0];
            if (!(c == '_' || Char.IsLetter(c)) || name.IndexOf(' ') != -1)
            {
                throw new ArgumentException("Invalid name: '" + name + "'; Names must begin with a letter or underscore and not contain spaces");
            }
        }

        public String RefersToFormula
        {
            get
            {
                if (_definedNameRec.IsFunctionName)
                {
                    throw new InvalidOperationException("Only applicable to named ranges");
                }
                Ptg[] ptgs = _definedNameRec.NameDefinition;
                if (ptgs.Length < 1)
                {
                    // 'refersToFormula' has not been set yet
                    return null;
                }
                return HSSFFormulaParser.ToFormulaString(book, ptgs);
            }
            set
            {
                Ptg[] ptgs = HSSFFormulaParser.Parse(value, book, NPOI.SS.Formula.FormulaType.NamedRange, SheetIndex);
                _definedNameRec.NameDefinition = ptgs;
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
                return _definedNameRec.SheetNumber - 1;
            }
            set 
            {
                int lastSheetIx = book.NumberOfSheets - 1;
                if (value < -1 || value > lastSheetIx)
                {
                    throw new ArgumentException("Sheet index (" + value + ") is out of range" +
                            (lastSheetIx == -1 ? "" : (" (0.." + lastSheetIx + ")")));
                }

                _definedNameRec.SheetNumber = (value + 1);
            }
        }
        public string Comment
        {
            get 
            {
                if (_commentRec != null)
                {
                    // Prefer the comment record if it has text in it
                    if (_commentRec.CommentText != null &&
                          _commentRec.CommentText.Length > 0)
                    {
                        return _commentRec.CommentText;
                    }
                }
                return _definedNameRec.DescriptionText; 
            }
            set { _definedNameRec.DescriptionText = value; }
        }


        //
        /// <summary>
        /// Sets the NameParsedFormula structure that specifies the formula for the defined name.
        /// </summary>
        /// <param name="ptgs">the sequence of {@link Ptg}s for the formula.</param>
        public void SetNameDefinition(Ptg[] ptgs)
        {
            _definedNameRec.NameDefinition = (ptgs);
        }

        /// <summary>
        /// Tests if this name points to a cell that no longer exists
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the name refers to a deleted cell; otherwise, <c>false</c>.
        /// </value>
        public bool IsDeleted
        {
            get
            {
                Ptg[] ptgs = _definedNameRec.NameDefinition;
                return Ptg.DoesFormulaReferToDeletedCell(ptgs);
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is function name.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is function name; otherwise, <c>false</c>.
        /// </value>
        public bool IsFunctionName
        {
            get
            {
                return _definedNameRec.IsFunctionName;
            }
        }
        /**
     * Indicates that the defined name refers to a user-defined function.
     * This attribute is used when there is an add-in or other code project associated with the file.
     *
     * @param value <c>true</c> indicates the name refers to a function.
     */
        public void SetFunction(bool value)
        {
            _definedNameRec.SetFunction(value);
        }
        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(_definedNameRec.NameText);
            sb.Append("]");
            return sb.ToString();
        }
    }
}