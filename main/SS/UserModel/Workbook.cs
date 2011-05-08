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


    /**
     * High level representation of a Excel workbook.  This is the first object most users
     * will construct whether they are Reading or writing a workbook.  It is also the
     * top level object for creating new sheets/etc.
     */
    public interface Workbook : IDisposable
    {


        /**
         * Convenience method to get the active sheet.  The active sheet is is the sheet
         * which is currently displayed when the workbook is viewed in Excel.
         * 'Selected' sheet(s) is a distinct concept.
         *
         * @return the index of the active sheet (0-based)
         */
        int ActiveSheetIndex { get; set; }

        /**
         * Gets the first tab that is displayed in the list of tabs in excel.
         *
         * @return the first tab that to display in the list of tabs (0-based).
         */
        int FirstVisibleTab { get; set; }

        /**
         * Sets the order of appearance for a given sheet.
         *
         * @param sheetname the name of the sheet to reorder
         * @param pos the position that we want to insert the sheet into (0 based)
         */
        void SetSheetOrder(String sheetname, int pos);

        /**
         * Sets the tab whose data is actually seen when the sheet is opened.
         * This may be different from the "selected sheet" since excel seems to
         * allow you to show the data of one sheet when another is seen "selected"
         * in the tabs (at the bottom).
         *
         * @see Sheet#SetSelected(bool)
         * @param index the index of the sheet to select (0 based)
         */
        void SetSelectedTab(int index);
        /**
         * Convenience method to set the active sheet.  The active sheet is is the sheet
         * which is currently displayed when the workbook is viewed in Excel.
         * 'Selected' sheet(s) is a distinct concept.
         *
         * @param sheetIndex index of the active sheet (0-based)
         */
        void SetActiveSheet(int sheetIndex);

        /**
         * Set the sheet name
         *
         * @param sheet sheet number (0 based)
         * @return Sheet name
         */
        String GetSheetName(int sheet);
        /**
         * Set the sheet name.
         *
         * @param sheet number (0 based)
         * @throws IllegalArgumentException if the name is greater than 31 chars or contains <code>/\?*[]</code>
         */
        void SetSheetName(int sheet, String name);
        /**
         * Returns the index of the sheet by his name
         *
         * @param name the sheet name
         * @return index of the sheet (0 based)
         */
        int GetSheetIndex(String name);

        /**
         * Returns the index of the given sheet
         *
         * @param sheet the sheet to look up
         * @return index of the sheet (0 based)
         */
        int GetSheetIndex(Sheet sheet);

        /**
         * Sreate an Sheet for this Workbook, Adds it to the sheets and returns
         * the high level representation.  Use this to create new sheets.
         *
         * @return Sheet representing the new sheet.
         */
        Sheet CreateSheet();

        /**
         * Create an Sheet for this Workbook, Adds it to the sheets and returns
         * the high level representation.  Use this to create new sheets.
         *
         * @param sheetname  sheetname to set for the sheet.
         * @return Sheet representing the new sheet.
         * @throws ArgumentException if the name is greater than 31 chars or Contains <code>/\?*[]</code>
         */
        Sheet CreateSheet(String sheetname);

        /**
         * Create an Sheet from an existing sheet in the Workbook.
         *
         * @return Sheet representing the Cloned sheet.
         */
        Sheet CloneSheet(int sheetNum);


        /**
         * Get the number of spreadsheets in the workbook
         *
         * @return the number of sheets
         */
        int NumberOfSheets { get; }

        /**
         * Get the Sheet object at the given index.
         *
         * @param index of the sheet number (0-based physical & logical)
         * @return Sheet at the provided index
         */
        Sheet GetSheetAt(int index);

        /**
         * Get sheet with the given name
         *
         * @param name of the sheet
         * @return Sheet with the name provided or <code>null</code> if it does not exist
         */
        Sheet GetSheet(String name);

        /**
         * Removes sheet at the given index
         *
         * @param index of the sheet to remove (0-based)
         */
        void RemoveSheetAt(int index);

        /**
         * Sets the repeating rows and columns for a sheet (as found in
         * File->PageSetup->Sheet).  This is function is included in the workbook
         * because it Creates/modifies name records which are stored at the
         * workbook level.
         * <p>
         * To set just repeating columns:
         * <pre>
         *  workbook.SetRepeatingRowsAndColumns(0,0,1,-1-1);
         * </pre>
         * To set just repeating rows:
         * <pre>
         *  workbook.SetRepeatingRowsAndColumns(0,-1,-1,0,4);
         * </pre>
         * To remove all repeating rows and columns for a sheet.
         * <pre>
         *  workbook.SetRepeatingRowsAndColumns(0,-1,-1,-1,-1);
         * </pre>
         *
         * @param sheetIndex    0 based index to sheet.
         * @param startColumn   0 based start of repeating columns.
         * @param endColumn     0 based end of repeating columns.
         * @param startRow      0 based start of repeating rows.
         * @param endRow        0 based end of repeating rows.
         */
        void SetRepeatingRowsAndColumns(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow);

        /**
         * Create a new Font and add it to the workbook's font table
         *
         * @return new font object
         */
        Font CreateFont();

        /**
         * Finds a font that matches the one with the supplied attributes
         *
         * @return the font with the matched attributes or <code>null</code>
         */
        Font FindFont(short boldWeight, short color, short fontHeight, String name, bool italic, bool strikeout, short typeOffset, byte underline);

        /**
         * Get the number of fonts in the font table
         *
         * @return number of fonts
         */
        short NumberOfFonts { get; }

        /**
         * Get the font at the given index number
         *
         * @param idx  index number (0-based)
         * @return font at the index
         */
        Font GetFontAt(short idx);

        /**
         * Create a new Cell style and add it to the workbook's style table
         *
         * @return the new Cell Style object
         */
        CellStyle CreateCellStyle();

        /**
         * Get the number of styles the workbook Contains
         *
         * @return count of cell styles
         */
        short NumCellStyles { get; }

        /**
         * Get the cell style object at the given index
         *
         * @param idx  index within the set of styles (0-based)
         * @return CellStyle object at the index
         */
        CellStyle GetCellStyleAt(short idx);

        /**
         * Write out this workbook to an OutPutstream.
         *
         * @param stream - the java OutPutStream you wish to write to
         * @exception IOException if anything can't be written.
         */
        void Write(Stream stream);

        /**
         * @return the total number of defined names in this workbook
         */
        int NumberOfNames { get; }

        /**
         * @param name the name of the defined name
         * @return the defined name with the specified name. <code>null</code> if not found.
         */
        Name GetName(String name);
        /**
         * @param nameIndex position of the named range (0-based)
         * @return the defined name at the specified index
         * @throws ArgumentException if the supplied index is invalid
         */
        Name GetNameAt(int nameIndex);

        /**
         * Creates a new (unInitialised) defined name in this workbook
         *
         * @return new defined name object
         */
        Name CreateName();

        /**
         * Gets the defined name index by name<br/>
         * <i>Note:</i> Excel defined names are case-insensitive and
         * this method performs a case-insensitive search.
         *
         * @param name the name of the defined name
         * @return zero based index of the defined name. <tt>-1</tt> if not found.
         */
        int GetNameIndex(String name);

        /**
         * Remove the defined name at the specified index
         *
         * @param index named range index (0 based)
         */
        void RemoveName(int index);

        /**
         * Remove a defined name by name
         *
          * @param name the name of the defined name
         */
        void RemoveName(String name);

        /**
        * Sets the printarea for the sheet provided
        * <p>
        * i.e. Reference = $A$1:$B$2
        * @param sheetIndex Zero-based sheet index (0 Represents the first sheet to keep consistent with java)
        * @param reference Valid name Reference for the Print Area
        */
        void SetPrintArea(int sheetIndex, String reference);

        /**
         * For the Convenience of Java Programmers maintaining pointers.
         * @see #SetPrintArea(int, String)
         * @param sheetIndex Zero-based sheet index (0 = First Sheet)
         * @param startColumn Column to begin printarea
         * @param endColumn Column to end the printarea
         * @param startRow Row to begin the printarea
         * @param endRow Row to end the printarea
         */
        void SetPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow);

        /**
         * Retrieves the reference for the printarea of the specified sheet,
         * the sheet name is Appended to the reference even if it was not specified.
         *
         * @param sheetIndex Zero-based sheet index (0 Represents the first sheet to keep consistent with java)
         * @return String Null if no print area has been defined
         */
        String GetPrintArea(int sheetIndex);

        /**
         * Delete the printarea for the sheet specified
         *
         * @param sheetIndex Zero-based sheet index (0 = First Sheet)
         */
        void RemovePrintArea(int sheetIndex);

        /**
         * Retrieves the current policy on what to do when
         *  Getting missing or blank cells from a row.
         * <p>
         * The default is to return blank and null cells.
         *  {@link MissingCellPolicy}
         * </p>
         */
        MissingCellPolicy MissingCellPolicy { get; set; }


        /**
         * Returns the instance of DataFormat for this workbook.
         *
         * @return the DataFormat object
         */
        DataFormat CreateDataFormat();

        /**
         * Adds a picture to the workbook.
         *
         * @param pictureData       The bytes of the picture
         * @param format            The format of the picture.
         *
         * @return the index to this picture (1 based).
         * @see #PICTURE_TYPE_EMF
         * @see #PICTURE_TYPE_WMF
         * @see #PICTURE_TYPE_PICT
         * @see #PICTURE_TYPE_JPEG
         * @see #PICTURE_TYPE_PNG
         * @see #PICTURE_TYPE_DIB
         */
        int AddPicture(byte[] pictureData, PictureType format);

        /**
         * Gets all pictures from the Workbook.
         *
         * @return the list of pictures (a list of {@link PictureData} objects.)
         */
        IList GetAllPictures();

        /**
         * Returns an object that handles instantiating concrete
         * classes of the various instances one needs for  HSSF and XSSF.
         */
        CreationHelper GetCreationHelper();

        /**
         * @return <code>false</code> if this workbook is not visible in the GUI
         */
        bool IsHidden { get; set; }
        /**
         * Check whether a sheet is hidden.
         * <p>
         * Note that a sheet could instead be set to be very hidden, which is different
         *  ({@link #isSheetVeryHidden(int)})
         * </p>
         * @param sheetIx Number
         * @return <code>true</code> if sheet is hidden
         */
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
        void SetSheetHidden(int sheetIx, bool hidden);

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

    }
}
