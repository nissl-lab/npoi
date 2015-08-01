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

namespace NPOI.SS.UserModel
{
    using System;
    using System.Collections;
    using System.IO;
    using NPOI.SS.Formula.Udf;
    using System.Collections.Generic;

    public enum SheetState : int
    {
        /// <summary>
        /// Indicates the sheet is visible.
        /// </summary>
        Visible = 0,

        /// <summary>
        /// Indicates the book window is hidden, but can be shown by the user via the user interface.
        /// </summary>
        Hidden = 1,

        /// <summary>
        /// Indicates the sheet is hidden and cannot be shown in the user interface (UI).
        /// </summary>
        /// <remarks>
        /// In Excel this state is only available programmatically in VBA:
        /// ThisWorkbook.Sheets("MySheetName").Visible = xlSheetVeryHidden
        /// 
        /// </remarks>
        VeryHidden = 2
    }

    /// <summary>
    /// High level interface of a Excel workbook.  This is the first object most users 
    /// will construct whether they are reading or writing a workbook.  It is also the
    /// top level object for creating new sheets/etc.
    /// This interface is shared between the implementation specific to xls and xlsx.
    /// This way it is possible to access Excel workbooks stored in both formats.
    /// </summary>
    public interface IWorkbook
    {

        /// <summary>
        /// get the active sheet.  The active sheet is is the sheet
        /// which is currently displayed when the workbook is viewed in Excel.
        /// </summary>
        int ActiveSheetIndex { get; }

        /// <summary>
        /// Gets the first tab that is displayed in the list of tabs in excel.
        /// </summary>
        int FirstVisibleTab { get; set; }

        /// <summary>
        /// Sets the order of appearance for a given sheet.
        /// </summary>
        /// <param name="sheetname">the name of the sheet to reorder</param>
        /// <param name="pos">the position that we want to insert the sheet into (0 based)</param>
        void SetSheetOrder(String sheetname, int pos);

        /// <summary>
        /// Sets the tab whose data is actually seen when the sheet is opened.
        /// This may be different from the "selected sheet" since excel seems to
        /// allow you to show the data of one sheet when another is seen "selected"
        /// in the tabs (at the bottom).
        /// </summary>
        /// <param name="index">the index of the sheet to select (0 based)</param>
        void SetSelectedTab(int index);

        /// <summary>
        /// set the active sheet. The active sheet is is the sheet
        /// which is currently displayed when the workbook is viewed in Excel.
        /// </summary>
        /// <param name="sheetIndex">index of the active sheet (0-based)</param>
        void SetActiveSheet(int sheetIndex);

        /// <summary>
        /// Set the sheet name
        /// </summary>
        /// <param name="sheet">sheet number (0 based)</param>
        /// <returns>Sheet name</returns>
        String GetSheetName(int sheet);

        /// <summary>
        /// Set the sheet name.
        /// </summary>
        /// <param name="sheet">sheet number (0 based)</param>
        /// <param name="name">sheet name</param>
        void SetSheetName(int sheet, String name);
   
        /// <summary>
        /// Returns the index of the sheet by its name
        /// </summary>
        /// <param name="name">the sheet name</param>
        /// <returns>index of the sheet (0 based)</returns>
        int GetSheetIndex(String name);

        /// <summary>
        /// Returns the index of the given sheet
        /// </summary>
        /// <param name="sheet">the sheet to look up</param>
        /// <returns>index of the sheet (0 based)</returns>
        int GetSheetIndex(ISheet sheet);

        /// <summary>
        /// Sreate an Sheet for this Workbook, Adds it to the sheets and returns
        /// the high level representation.  Use this to create new sheets.
        /// </summary>
        /// <returns></returns>
        ISheet CreateSheet();

        /// <summary>
        /// Create an Sheet for this Workbook, Adds it to the sheets and returns
        /// the high level representation.  Use this to create new sheets.
        /// </summary>
        /// <param name="sheetname">sheetname to set for the sheet.</param>
        /// <returns>Sheet representing the new sheet.</returns>
        ISheet CreateSheet(String sheetname);

        /// <summary>
        /// Create an Sheet from an existing sheet in the Workbook.
        /// </summary>
        /// <param name="sheetNum"></param>
        /// <returns></returns>
        ISheet CloneSheet(int sheetNum);

         /// <summary>
         /// Get the number of spreadsheets in the workbook
         /// </summary>
        int NumberOfSheets { get; }

        /// <summary>
        /// Get the Sheet object at the given index.
        /// </summary>
        /// <param name="index">index of the sheet number (0-based physical &amp; logical)</param>
        /// <returns>Sheet at the provided index</returns>
        ISheet GetSheetAt(int index);

        /// <summary>
        /// Get sheet with the given name
        /// </summary>
        /// <param name="name">name of the sheet</param>
        /// <returns>Sheet with the name provided or null if it does not exist</returns>
        ISheet GetSheet(String name);

        /// <summary>
        /// Removes sheet at the given index
        /// </summary>
        /// <param name="index"></param>
        void RemoveSheetAt(int index);

        /// <summary>
        /// Enumerate sheets
        /// </summary>
        /// <returns></returns>
        IEnumerator GetEnumerator();
        /**
         * To set just repeating columns:
         *  workbook.SetRepeatingRowsAndColumns(0,0,1,-1-1);
         * To set just repeating rows:
         *  workbook.SetRepeatingRowsAndColumns(0,-1,-1,0,4);
         * To remove all repeating rows and columns for a sheet.
         *  workbook.SetRepeatingRowsAndColumns(0,-1,-1,-1,-1);
         */
        /// <summary>
        /// Sets the repeating rows and columns for a sheet (as found in
        /// File->PageSetup->Sheet).  This is function is included in the workbook
        /// because it Creates/modifies name records which are stored at the
        /// workbook level.
        /// </summary>
        /// <param name="sheetIndex">0 based index to sheet.</param>
        /// <param name="startColumn">0 based start of repeating columns.</param>
        /// <param name="endColumn">0 based end of repeating columns.</param>
        /// <param name="startRow">0 based start of repeating rows.</param>
        /// <param name="endRow">0 based end of repeating rows.</param>
        [Obsolete("use Sheet#setRepeatingRows(CellRangeAddress) or Sheet#setRepeatingColumns(CellRangeAddress)")]
        void SetRepeatingRowsAndColumns(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow);


        /// <summary>
        /// Create a new Font and add it to the workbook's font table
        /// </summary>
        /// <returns></returns>
        IFont CreateFont();

        /// <summary>
        /// Finds a font that matches the one with the supplied attributes
        /// </summary>
        /// <param name="boldWeight"></param>
        /// <param name="color"></param>
        /// <param name="fontHeight"></param>
        /// <param name="name"></param>
        /// <param name="italic"></param>
        /// <param name="strikeout"></param>
        /// <param name="typeOffset"></param>
        /// <param name="underline"></param>
        /// <returns>the font with the matched attributes or null</returns>
        IFont FindFont(short boldWeight, short color, short fontHeight, String name, bool italic, bool strikeout, FontSuperScript typeOffset, FontUnderlineType underline);

        /// <summary>
        /// Get the number of fonts in the font table
        /// </summary>
        short NumberOfFonts { get; }

        /// <summary>
        /// Get the font at the given index number
        /// </summary>
        /// <param name="idx">index number (0-based)</param>
        /// <returns>font at the index</returns>
        IFont GetFontAt(short idx);

        /// <summary>
        /// Create a new Cell style and add it to the workbook's style table
        /// </summary>
        /// <returns>the new Cell Style object</returns>
        ICellStyle CreateCellStyle();

        /// <summary>
        /// Get the number of styles the workbook Contains
        /// </summary>
        short NumCellStyles { get; }

        /// <summary>
        /// Get the cell style object at the given index
        /// </summary>
        /// <param name="idx">index within the set of styles (0-based)</param>
        /// <returns>CellStyle object at the index</returns>
        ICellStyle GetCellStyleAt(short idx);

        /// <summary>
        /// Write out this workbook to an OutPutstream.
        /// </summary>
        /// <param name="stream">the stream you wish to write to</param>
        void Write(Stream stream);

        /// <summary>
        /// the total number of defined names in this workbook
        /// </summary>
        int NumberOfNames { get; }

        /// <summary>
        /// the defined name with the specified name.
        /// </summary>
        /// <param name="name">the name of the defined name</param>
        /// <returns>the defined name with the specified name. null if not found</returns>
        IName GetName(String name);

        /// <summary>
        /// the defined name at the specified index
        /// </summary>
        /// <param name="nameIndex">position of the named range (0-based)</param>
        /// <returns></returns>
        IName GetNameAt(int nameIndex);

        /// <summary>
        /// Creates a new (unInitialised) defined name in this workbook
        /// </summary>
        /// <returns>new defined name object</returns>
        IName CreateName();

        /// <summary>
        /// Gets the defined name index by name
        /// </summary>
        /// <param name="name">the name of the defined name</param>
        /// <returns>zero based index of the defined name.</returns>
        int GetNameIndex(String name);

        /// <summary>
        /// Remove the defined name at the specified index
        /// </summary>
        /// <param name="index">named range index (0 based)</param>
        void RemoveName(int index);

        /// <summary>
        /// Remove a defined name by name
        /// </summary>
        /// <param name="name">the name of the defined name</param>
        void RemoveName(String name);

        /// <summary>
        /// Adds the linking required to allow formulas referencing the specified 
        /// external workbook to be added to this one. In order for formulas 
        /// such as "[MyOtherWorkbook]Sheet3!$A$5" to be added to the file, 
        /// some linking information must first be recorded. Once a given external 
        /// workbook has been linked, then formulas using it can added. Each workbook 
        /// needs linking only once. <br/>
        /// This linking only applies for writing formulas. 
        /// To link things for evaluation, see {@link FormulaEvaluator#setupReferencedWorkbooks(java.util.Map)}
        /// </summary>
        /// <param name="name">The name the workbook will be referenced as in formulas</param>
        /// <param name="workbook">The open workbook to fetch the link required information from</param>
        /// <returns></returns>
        int LinkExternalWorkbook(String name, IWorkbook workbook);
    

        /// <summary>
        /// Sets the printarea for the sheet provided
        /// </summary>
        /// <param name="sheetIndex">Zero-based sheet index</param>
        /// <param name="reference">Valid name Reference for the Print Area</param>
        void SetPrintArea(int sheetIndex, String reference);


        /// <summary>
        /// Sets the printarea for the sheet provided
        /// </summary>
        /// <param name="sheetIndex">Zero-based sheet index (0 = First Sheet)</param>
        /// <param name="startColumn">Column to begin printarea</param>
        /// <param name="endColumn">Column to end the printarea</param>
        /// <param name="startRow">Row to begin the printarea</param>
        /// <param name="endRow">Row to end the printarea</param>
        void SetPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow);

        /// <summary>
        /// Retrieves the reference for the printarea of the specified sheet, 
        /// the sheet name is Appended to the reference even if it was not specified.
        /// </summary>
        /// <param name="sheetIndex">Zero-based sheet index</param>
        /// <returns>Null if no print area has been defined</returns>
        String GetPrintArea(int sheetIndex);

        /// <summary>
        /// Delete the printarea for the sheet specified
        /// </summary>
        /// <param name="sheetIndex">Zero-based sheet index (0 = First Sheet)</param>
        void RemovePrintArea(int sheetIndex);

        /// <summary>
        /// Retrieves the current policy on what to do when getting missing or blank cells from a row.
        /// </summary>
        MissingCellPolicy MissingCellPolicy { get; set; }

        /// <summary>
        /// Returns the instance of DataFormat for this workbook.
        /// </summary>
        /// <returns>the DataFormat object</returns>
        IDataFormat CreateDataFormat();

        /// <summary>
        /// Adds a picture to the workbook.
        /// </summary>
        /// <param name="pictureData">The bytes of the picture</param>
        /// <param name="format">The format of the picture.</param>
        /// <returns>the index to this picture (1 based).</returns>
        int AddPicture(byte[] pictureData, PictureType format);

        /// <summary>
        /// Gets all pictures from the Workbook.
        /// </summary>
        /// <returns>the list of pictures (a list of link PictureData objects.)</returns>
        IList GetAllPictures();

        /// <summary>
        /// Return an object that handles instantiating concrete classes of 
        /// the various instances one needs for  HSSF and XSSF.
        /// </summary>
        /// <returns></returns>
        ICreationHelper GetCreationHelper();

        /// <summary>
        /// if this workbook is not visible in the GUI
        /// </summary>
        bool IsHidden { get; set; }

        /// <summary>
        /// Check whether a sheet is hidden.
        /// </summary>
        /// <param name="sheetIx">number of sheet</param>
        /// <returns>true if sheet is hidden</returns>
        bool IsSheetHidden(int sheetIx);

        /**
         * Check whether a sheet is very hidden.
         * <p>
         * This is different from the normal hidden status
         *  ({@link #isSheetHidden(int)})
         * </p>
         * @param sheetIx sheet index to check
         * @return <code>true</code> if sheet is very hidden
         */
        bool IsSheetVeryHidden(int sheetIx);

        /**
         * Hide or unhide a sheet
         *
         * @param sheetIx the sheet index (0-based)
         * @param hidden True to mark the sheet as hidden, false otherwise
         */
        void SetSheetHidden(int sheetIx, SheetState hidden);

        /**
         * Hide or unhide a sheet.
         * <pre>
         *  0 = not hidden
         *  1 = hidden
         *  2 = very hidden.
         * </pre>
         * @param sheetIx The sheet number
         * @param hidden 0 for not hidden, 1 for hidden, 2 for very hidden
         */
        void SetSheetHidden(int sheetIx, int hidden);

        /// <summary>
        /// Register a new toolpack in this workbook.
        /// </summary>
        /// <param name="toopack">the toolpack to register</param>
        void AddToolPack(UDFFinder toopack);

        void Close();
    }
}
