﻿/* ====================================================================
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
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text; 
using Cysharp.Text;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.OpenXml4Net.Util;
using NPOI.SS;
using NPOI.SS.Formula.UDF;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    /**
     * Streaming version of XSSFWorkbook implementing the "BigGridDemo" strategy.
     *
     * This allows to write very large files without running out of memory as only
     * a configurable portion of the rows are kept in memory at any one time.
     *
     * You can provide a template workbook which is used as basis for the written
     * data.
     *
     * See https://poi.apache.org/spreadsheet/how-to.html#sxssf for details.
     *
     * Please note that there are still things that still may consume a large
     * amount of memory based on which features you are using, e.g. merged regions,
     * comments, ... are still only stored in memory and thus may require a lot of
     * memory if used extensively.
     *
     * SXSSFWorkbook defaults to using inline strings instead of a shared strings
     * table. This is very efficient, since no document content needs to be kept in
     * memory, but is also known to produce documents that are incompatible with
     * some clients. With shared strings enabled all unique strings in the document
     * has to be kept in memory. Depending on your document content this could use
     * a lot more resources than with shared strings disabled.
     *
     * Carefully review your memory budget and compatibility needs before deciding
     * whether to enable shared strings or not.
     */
    /// <summary>
    /// Streaming version of the XSSFWorkbook, originally implemented in the "BigGridDemo".
    /// </summary>
    public class SXSSFWorkbook : IWorkbook
    {
        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(SXSSFWorkbook));

        public const int DEFAULT_WINDOW_SIZE = 100;
        private readonly XSSFWorkbook _wb;
        public XSSFWorkbook XssfWorkbook
        {
            get { return _wb; }
        }

        private readonly Dictionary<SXSSFSheet, XSSFSheet> _sxFromXHash = new Dictionary<SXSSFSheet, XSSFSheet>();
        private readonly Dictionary<XSSFSheet, SXSSFSheet> _xFromSxHash = new Dictionary<XSSFSheet, SXSSFSheet>();

        private int _randomAccessWindowSize = DEFAULT_WINDOW_SIZE;

        /**
        * See the constructors for a more detailed description of the sliding window of rows.
        *
        * @return The number of rows that are kept in memory at once before flushing them out.
        */
        public int RandomAccessWindowSize
        {
            get { return _randomAccessWindowSize; }
            set
            {
                if (value <= 0)
                {
                    throw new ArgumentException("rowAccessWindowSize must be greater than 0 or -1");
                }
                _randomAccessWindowSize = value;
            }

        }

        /// <summary>
        /// whether temp file should be compresss.
        /// </summary>
        private bool _compressTmpFiles = false;

        /// <summary>
        /// setting this flag On allows to write large files;
        /// however, this can lead to compatibility issues
        /// </summary>
        private UseZip64 _useZip64 = UseZip64.Off;

        /// <summary>
        /// shared string table - a cache of strings in this workbook.
        /// </summary>
        private readonly SharedStringsTable _sharedStringSource;

        public int ActiveSheetIndex
        {
            get { return XssfWorkbook.ActiveSheetIndex; }
        }

        public int FirstVisibleTab
        {
            get { return XssfWorkbook.FirstVisibleTab; }

            set { XssfWorkbook.FirstVisibleTab = value; }
        }

        public int NumberOfSheets
        {
            get { return XssfWorkbook.NumberOfSheets; }
        }

        public short NumberOfFonts
        {
            get { return XssfWorkbook.NumberOfFonts; }
        }

        public int NumCellStyles
        {
            get { return XssfWorkbook.NumCellStyles; }
        }

        public int NumberOfNames
        {
            get { return XssfWorkbook.NumberOfNames; }
        }

        public MissingCellPolicy MissingCellPolicy
        {
            get { return XssfWorkbook.MissingCellPolicy; }

            set { XssfWorkbook.MissingCellPolicy = value; }
        }

        public bool IsHidden
        {
            get { return XssfWorkbook.IsHidden; }

            set { XssfWorkbook.IsHidden = value; }
        }

        public UseZip64 UseZip64
        {
            get { return _useZip64; }

            set { _useZip64 = value; }
        }


        #region Constructors
        /**
         * Construct an empty workbook and specify the window for row access.
         * <p>
         * When a new node is created via createRow() and the total number
         * of unflushed records would exceed the specified value, then the
         * row with the lowest index value is flushed and cannot be accessed
         * via getRow() anymore.
         * </p>
         * <p>
         * A value of -1 indicates unlimited access. In this case all
         * records that have not been flushed by a call to flush() are available
         * for random access.
         * </p>
         * <p>
         * A value of 0 is not allowed because it would flush any newly created row
         * without having a chance to specify any cells.
         * </p>
         *
         * @param rowAccessWindowSize the number of rows that are kept in memory until flushed out, see above.
         */
        public SXSSFWorkbook(int rowAccessWindowSize)
            : this(null /*workbook*/, rowAccessWindowSize)
        {
            
        }
        /// <summary>
        /// Construct a new workbook with default row window size
        /// </summary>
        public SXSSFWorkbook() : this(null)
        {

        }
        /**
         * Construct a workbook from a template.
         * <p>
         * There are three use-cases to use SXSSFWorkbook(XSSFWorkbook) :
         * <ol>
         *   <li>
         *       Append new sheets to existing workbooks. You can open existing
         *       workbook from a file or create on the fly with XSSF.
         *   </li>
         *   <li>
         *       Append rows to existing sheets. The row number MUST be greater
         *       than max(rownum) in the template sheet.
         *   </li>
         *   <li>
         *       Use existing workbook as a template and re-use global objects such
         *       as cell styles, formats, images, etc.
         *   </li>
         * </ol>
         * All three use cases can work in a combination.
         * </p>
         * What is not supported:
         * <ul>
         *   <li>
         *   Access initial cells and rows in the template. After constructing
         *   SXSSFWorkbook(XSSFWorkbook) all internal windows are empty and
         *   SXSSFSheet@getRow and SXSSFRow#getCell return null.
         *   </li>
         *   <li>
         *    Override existing cells and rows. The API silently allows that but
         *    the output file is invalid and Excel cannot read it.
         *   </li>
         * </ul>
         *
         * @param workbook  the template workbook
         */
        public SXSSFWorkbook(XSSFWorkbook workbook) 
            : this(workbook, DEFAULT_WINDOW_SIZE)
        {

        }
        /**
         * Constructs an workbook from an existing workbook.
         * <p>
         * When a new node is created via createRow() and the total number
         * of unflushed records would exceed the specified value, then the
         * row with the lowest index value is flushed and cannot be accessed
         * via getRow() anymore.
         * </p>
         * <p>
         * A value of -1 indicates unlimited access. In this case all
         * records that have not been flushed by a call to flush() are available
         * for random access.
         * </p>
         * <p>
         * A value of 0 is not allowed because it would flush any newly created row
         * without having a chance to specify any cells.
         * </p>
         *
         * @param rowAccessWindowSize the number of rows that are kept in memory until flushed out, see above.
         */
        public SXSSFWorkbook(XSSFWorkbook workbook, int rowAccessWindowSize)
            : this(workbook, rowAccessWindowSize, false)
        {

        }
        /**
         * Constructs an workbook from an existing workbook.
         * <p>
         * When a new node is created via createRow() and the total number
         * of unflushed records would exceed the specified value, then the
         * row with the lowest index value is flushed and cannot be accessed
         * via getRow() anymore.
         * </p>
         * <p>
         * A value of -1 indicates unlimited access. In this case all
         * records that have not been flushed by a call to flush() are available
         * for random access.
         * </p>
         * <p>
         * A value of 0 is not allowed because it would flush any newly created row
         * without having a chance to specify any cells.
         * </p>
         *
         * @param rowAccessWindowSize the number of rows that are kept in memory until flushed out, see above.
         * @param compressTmpFiles whether to use gzip compression for temporary files
         */
        public SXSSFWorkbook(XSSFWorkbook workbook, int rowAccessWindowSize, bool compressTmpFiles)
            : this(workbook, rowAccessWindowSize, compressTmpFiles, false)
        {

        }
        /**
         * Constructs an workbook from an existing workbook.
         * <p>
         * When a new node is created via createRow() and the total number
         * of unflushed records would exceed the specified value, then the
         * row with the lowest index value is flushed and cannot be accessed
         * via getRow() anymore.
         * </p>
         * <p>
         * A value of -1 indicates unlimited access. In this case all
         * records that have not been flushed by a call to flush() are available
         * for random access.
         * </p>
         * <p>
         * A value of 0 is not allowed because it would flush any newly created row
         * without having a chance to specify any cells.
         * </p>
         *
         * @param workbook  the template workbook
         * @param rowAccessWindowSize the number of rows that are kept in memory until flushed out, see above.
         * @param compressTmpFiles whether to use gzip compression for temporary files
         * @param useSharedStringsTable whether to use a shared strings table
         */
        /// <summary>
        /// Currently only supports writing not reading. E.g. the number of rows returned from a worksheet will be wrong etc.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="rowAccessWindowSize"></param>
        /// <param name="compressTmpFiles"></param>
        /// <param name="useSharedStringsTable"></param>
        public SXSSFWorkbook(XSSFWorkbook workbook, int rowAccessWindowSize, bool compressTmpFiles, bool useSharedStringsTable)
        {
            RandomAccessWindowSize = rowAccessWindowSize;

            _compressTmpFiles = compressTmpFiles;

            if (workbook == null)
            {
                _wb = new XSSFWorkbook();
                _sharedStringSource = useSharedStringsTable ? XssfWorkbook.GetSharedStringSource() : null;
            }
            else
            {
                _wb = workbook;
                _sharedStringSource = useSharedStringsTable ? XssfWorkbook.GetSharedStringSource() : null;
                var numberOfSheets = XssfWorkbook.NumberOfSheets;
                for (int i = 0; i < numberOfSheets; i++)
                {
                    XSSFSheet sheet = (XSSFSheet)XssfWorkbook.GetSheetAt(i);
                    CreateAndRegisterSXSSFSheet(sheet);
                }
            }
        }

        #endregion

        private SXSSFSheet CreateAndRegisterSXSSFSheet(ISheet xSheet)
        {
            SXSSFSheet sxSheet;
            try
            {
                sxSheet = new SXSSFSheet(this, (XSSFSheet)xSheet);
            }
            catch (IOException ioe)
            {
                throw new RuntimeException(ioe);
            }
            RegisterSheetMapping(sxSheet, (XSSFSheet)xSheet);
            return sxSheet;
        }

        private void RegisterSheetMapping(SXSSFSheet sxSheet, XSSFSheet xSheet)
        {
            _sxFromXHash.Add(sxSheet, xSheet);
            _xFromSxHash.Add(xSheet, sxSheet);
        }

        private void DeregisterSheetMapping(XSSFSheet xSheet)
        {
            SXSSFSheet sxSheet = GetSXSSFSheet(xSheet);

            try
            {
                sxSheet.SheetDataWriter.Close();
            }
            catch (IOException)
            {
                // ignore exception here
            }

            _sxFromXHash.Remove(sxSheet);

            _xFromSxHash.Remove(xSheet);
        }


        public XSSFSheet GetXSSFSheet(SXSSFSheet sheet)
        {
            if (sheet != null && _sxFromXHash.TryGetValue(sheet, out XSSFSheet xssfSheet))
                return xssfSheet;
            else
                return null;
        }

        public SXSSFSheet GetSXSSFSheet(XSSFSheet sheet)
        {
            if (sheet != null && _xFromSxHash.TryGetValue(sheet, out SXSSFSheet sxssfSheet))
                return sxssfSheet;
            else
                return null;
        }

        /**
         * Set whether temp files should be compressed.
         * <p>
         *   SXSSF writes sheet data in temporary files (a temp file per-sheet)
         *   and the size of these temp files can grow to to a very large size,
         *   e.g. for a 20 MB csv data the size of the temp xml file become few GB large.
         *   If the "compress" flag is set to <code>true</code> then the temporary XML is gzipped.
         * </p>
         * <p>
         *     Please note the the "compress" option may cause performance penalty.
         * </p>
         * @param compress whether to compress temp files
         */
        public bool CompressTempFiles
        {
            get
            {
                return _compressTmpFiles;
            }
            set
            {
                _compressTmpFiles = value;
            }
            
        }

        public SheetDataWriter CreateSheetDataWriter()
        {
            if (_compressTmpFiles)
            {
                return new GZIPSheetDataWriter(_sharedStringSource);
            }

            return new SheetDataWriter(_sharedStringSource);
        }

        private XSSFSheet GetSheetFromZipEntryName(string sheetRef)
        {
            foreach (XSSFSheet sheet in _sxFromXHash.Values)
            {
                if (sheetRef.Equals(sheet.GetPackagePart().PartName.Name.Substring(1))) return sheet;
            }
            return null;
        }

        private void InjectData(FileInfo zipfile, Stream outStream, bool leaveOpen)
        {
            // don't use ZipHelper.openZipFile here - see #59743
            ZipFile zip = new ZipFile(zipfile.FullName);
            try
            {
                ZipOutputStream zos = new ZipOutputStream(outStream);
                try
                {
                    zos.IsStreamOwner = !leaveOpen;
                    zos.UseZip64 = _useZip64;
                    //ZipEntrySource zipEntrySource = new ZipFileZipEntrySource(zip);
                    //var en =  zipEntrySource.Entries;
                    var en = zip.GetEnumerator();
                    while (en.MoveNext())
                    {
                        var ze = (ZipEntry)en.Current;
                        zos.PutNextEntry(new ZipEntry(ze.Name));
                        var inputStream = zip.GetInputStream(ze);
                        XSSFSheet xSheet = GetSheetFromZipEntryName(ze.Name);
                        if (xSheet != null)
                        {
                            SXSSFSheet sxSheet = GetSXSSFSheet(xSheet);
                            var xis = sxSheet.GetWorksheetXMLInputStream();
                            try
                            {
                                CopyStreamAndInjectWorksheet(inputStream, zos, xis);
                            }
                            finally
                            {
                                xis.Close();
                            }
                        }
                        else
                        {
                            inputStream.CopyTo(zos);
                        }
                        inputStream.Close();
                    }
                }
                finally
                {
                    zos.Close();
                }
            }
            finally
            {
                zip.Close();
            }
        }

        private static void CopyStreamAndInjectWorksheet(Stream inputStream, Stream outputStream, Stream worksheetData)
        {
            StreamReader inReader = new StreamReader(inputStream, Encoding.UTF8);
            StreamWriter outWriter = new StreamWriter(outputStream, Encoding.UTF8);
            bool needsStartTag = true;
            int c;
            int pos = 0;
            String s = "<sheetData";

            StringBuilder sb = new StringBuilder();
            int n = s.Length;
            //Copy from "in" to "out" up to the string "<sheetData/>" or "</sheetData>" (excluding).
            while (((c = inReader.Read()) != -1))
            {
                if ((char)c == (char)s[pos])
                {
                    pos++;
                    if (pos == n)
                    {
                        if ("<sheetData".Equals(s))
                        {
                            c = inReader.Read();
                            if (c == -1)
                            {
                                outWriter.Write(s);
                                sb.Append(s);
                                break;
                            }
                            if ((char)c == '>')
                            {
                                // Found <sheetData>
                                outWriter.Write(s);
                                sb.Append(s);
                                outWriter.Write((char)c);
                                sb.Append((char) c);
                                s = "</sheetData>";
                                n = s.Length;
                                pos = 0;
                                needsStartTag = false;
                                continue;
                            }
                            if ((char)c == '/')
                            {
                                // Found <sheetData/
                                c = inReader.Read();
                                if (c == -1)
                                {
                                    outWriter.Write(s);
                                    sb.Append(s);
                                    break;
                                }
                                if ((char)c == '>')
                                {
                                    // Found <sheetData/>
                                    break;
                                }

                                outWriter.Write(s);
                                sb.Append(s);
                                outWriter.Write('/');
                                sb.Append('/');
                                outWriter.Write((char)c);
                                sb.Append((char) c);
                                pos = 0;
                                continue;
                            }

                            outWriter.Write(s);
                            sb.Append(s);
                            outWriter.Write('/');
                            sb.Append('/');
                            outWriter.Write((char)c);
                            sb.Append((char)c);
                            pos = 0;
                            continue;
                        }
                        else
                        {
                            // Found </sheetData>
                            break;
                        }
                    }
                }
                else
                {
                    if (pos > 0)
                    {
                        outWriter.Write(s.Substring(0,pos));
                        sb.Append(s, 0, pos);
                    }
                    if (c == s[0])
                    {
                        pos = 1;
                    }
                    else
                    {
                        outWriter.Write((char)c);
                        sb.Append((char) c);
                        pos = 0;
                    }
                }
            }
            outWriter.Flush();
            if (needsStartTag)
            {
                outWriter.Write("<sheetData>\n");
                sb.Append("<sheetData>\n");
                outWriter.Flush();
            }
            //Copy the worksheet data to "out".
            worksheetData.CopyTo(outputStream);
            outWriter.Write("</sheetData>");
            outWriter.Flush();
            //Copy the rest of "in" to "out".
            while (((c = inReader.Read()) != -1))
            {
                outWriter.Write((char) c);
                sb.Append((char)c);
            }
            outWriter.Flush();
            
        }

        public void SetSheetOrder(string sheetname, int pos)
        {
            XssfWorkbook.SetSheetOrder(sheetname, pos);
        }

        public void SetSelectedTab(int index)
        {
            XssfWorkbook.SetSelectedTab(index);
        }

        public void SetActiveSheet(int sheetIndex)
        {
            XssfWorkbook.SetActiveSheet(sheetIndex);
        }

        public string GetSheetName(int sheet)
        {
            return XssfWorkbook.GetSheetName(sheet);
        }

        public void SetSheetName(int sheet, string name)
        {
            XssfWorkbook.SetSheetName(sheet, name);
        }

        public int GetSheetIndex(string name)
        {
            return XssfWorkbook.GetSheetIndex(name);
        }

        public int GetSheetIndex(ISheet sheet)
        {
            return XssfWorkbook.GetSheetIndex(GetXSSFSheet((SXSSFSheet)sheet));
        }

        public ISheet CreateSheet()
        {
            return CreateAndRegisterSXSSFSheet(XssfWorkbook.CreateSheet());
        }

        public ISheet CreateSheet(string sheetname)
        {
            SS.Util.WorkbookUtil.ValidateSheetName(sheetname);
            return CreateAndRegisterSXSSFSheet(XssfWorkbook.CreateSheet(sheetname));
        }

        public ISheet CloneSheet(int sheetNum)
        {
            throw new RuntimeException("NotImplemented");
        }

        public ISheet GetSheetAt(int index)
        {
            return GetSXSSFSheet((XSSFSheet)XssfWorkbook.GetSheetAt(index));
        }

        public ISheet GetSheet(string name)
        {
            return GetSXSSFSheet((XSSFSheet)XssfWorkbook.GetSheet(name));
        }

        public void RemoveSheetAt(int index)
        {
            // Get the sheet to be removed
            var xSheet = (XSSFSheet)XssfWorkbook.GetSheetAt(index);
            var sxSheet = GetSXSSFSheet(xSheet);

            // De-register it
            XssfWorkbook.RemoveSheetAt(index);
            DeregisterSheetMapping(xSheet);

            // Clean up temporary resources
            try
            {
                sxSheet.Dispose();
            }
            catch (IOException e)
            {
                logger.Log(POILogger.WARN, e);
            }
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            return new SheetEnumerator<SXSSFSheet>(XssfWorkbook, this);
        }

        public IFont CreateFont()
        {
            return XssfWorkbook.CreateFont();
        }

        [Obsolete("deprecated in poi 3.16")]
        public IFont FindFont(short boldWeight, short color, short fontHeight, string name, bool italic, bool strikeout, FontSuperScript typeOffset, FontUnderlineType underline)
        {
            return XssfWorkbook.FindFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
        }

        /**
         * Finds a font that matches the one with the supplied attributes
         *
         * @return the font with the matched attributes or <code>null</code>
         */
        public IFont FindFont(bool bold, short color, short fontHeight, String name, bool italic, bool strikeout, FontSuperScript typeOffset, FontUnderlineType underline)
        { 
            return XssfWorkbook.FindFont(bold, color, fontHeight, name, italic, strikeout, typeOffset, underline);
        }

        public IFont GetFontAt(short idx)
        {
            return XssfWorkbook.GetFontAt(idx);
        }

        public ICellStyle CreateCellStyle()
        {
            return XssfWorkbook.CreateCellStyle();
        }

        public ICellStyle GetCellStyleAt(int idx)
        {
            return XssfWorkbook.GetCellStyleAt(idx);
        }


        public void Close()
        {
            // ensure that any lingering writer is closed
            foreach (SXSSFSheet sheet in _xFromSxHash.Values)
            {
                try
                {
                    sheet.SheetDataWriter.Close();
                }
                catch (IOException e)
                {
                    logger.Log(POILogger.WARN,
                            "An exception occurred while closing sheet data writer for sheet "
                            + sheet.SheetName + ".", e);
                }
            }


            // Tell the base workbook to close, does nothing if 
            //  it's a newly created one
            XssfWorkbook.Close();
        }
        public void Write(Stream stream, bool leaveOpen = false)
        {
            FlushSheets();

            //Save the template
            var tmplFile = TempFile.CreateTempFile("poi-sxssf-template", ".xlsx");
            try
            {
                var os = new FileStream(tmplFile.FullName, FileMode.Open, FileAccess.ReadWrite);
                try
                {
                    XssfWorkbook.Write(os);
                }
                finally
                {
                    os.Close();
                }

                //Substitute the template entries with the generated sheet data files
                
                InjectData(tmplFile, stream, leaveOpen);
            }
            finally
            {
                tmplFile.Delete();
                if (File.Exists(tmplFile.FullName))
                {
                    throw new IOException("Could not delete temporary file after processing: " + tmplFile);
                }
            }
        }

        private void FlushSheets()
        {
            foreach (SXSSFSheet sheet in _xFromSxHash.Values)
            {
                sheet.FlushRows();
            }
        }

        /**
        * Dispose of temporary files backing this workbook on disk.
        * Calling this method will render the workbook unusable.
        * @return true if all temporary files were deleted successfully.
        */
        public bool Dispose()
        {
            var success = true;
            foreach (SXSSFSheet sheet in _sxFromXHash.Keys)
            {
                try
                {
                    success = sheet.Dispose() && success;
                }
                catch (IOException e)
                {
                    logger.Log(POILogger.WARN, e);
                    success = false;
                }
            }
            return success;
        }

        public IName GetName(string name)
        {
            return XssfWorkbook.GetName(name);
        }

        /**
         * Returns all defined names with the given name.
         *
         * @param name the name of the defined name
         * @return a list of the defined names with the specified name. An empty list is returned if none is found.
         */
        public IList<IName> GetNames(String name)
        {
            return XssfWorkbook.GetNames(name);
        }

        /// <summary>
        /// Returns all defined names
        /// </summary>
        /// <returns>Returns all defined names</returns>
        public IList<IName> GetAllNames()
        {
            return _wb.GetAllNames();
        }
        [Obsolete("Deprecated 3.16, New projects should avoid accessing named ranges by index.")]
        public IName GetNameAt(int nameIndex)
        {
            return XssfWorkbook.GetNameAt(nameIndex);
        }

        public IName CreateName()
        {
            return XssfWorkbook.CreateName();
        }

        [Obsolete("deprecated in 3.16 New projects should avoid accessing named ranges by index. GetName(String)} instead.")]
        public int GetNameIndex(string name)
        {
            return XssfWorkbook.GetNameIndex(name);
        }

        [Obsolete("deprecated in 3.16 New projects should use RemoveName(Name)")]
        public void RemoveName(int index)
        {
            XssfWorkbook.RemoveName(index);
        }

        [Obsolete("deprecated in 3.16 New projects should use RemoveName(IName Name)")]
        public void RemoveName(string name)
        {
            XssfWorkbook.RemoveName(name);
        }
        /// <summary>
        /// Remove the given defined name
        /// </summary>
        /// <param name="name">the name to remove</param>
        public void RemoveName(IName name)
        {
            _wb.RemoveName(name);
        }
        public int LinkExternalWorkbook(string name, IWorkbook workbook)
        {
            throw new NotImplementedException();
        }

        public void SetPrintArea(int sheetIndex, string reference)
        {
            XssfWorkbook.SetPrintArea(sheetIndex, reference);
        }

        public void SetPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow)
        {
            XssfWorkbook.SetPrintArea(sheetIndex, startColumn, endColumn, startRow, endRow);
        }

        public string GetPrintArea(int sheetIndex)
        {
            return XssfWorkbook.GetPrintArea(sheetIndex);
        }

        public void RemovePrintArea(int sheetIndex)
        {
            XssfWorkbook.RemovePrintArea(sheetIndex);
        }


        public IDataFormat CreateDataFormat()
        {
            return XssfWorkbook.CreateDataFormat();
        }

        public int AddPicture(byte[] pictureData, PictureType format)
        {
            return XssfWorkbook.AddPicture(pictureData, format);
        }

        public IList GetAllPictures()
        {
            return XssfWorkbook.GetAllPictures();
        }

        public ICreationHelper GetCreationHelper()
        {
            return new SXSSFCreationHelper(this);
        }
        public bool IsSheetHidden(int sheetIx)
        {
            return XssfWorkbook.IsSheetHidden(sheetIx);
        }

        public bool IsSheetVeryHidden(int sheetIx)
        {
            return XssfWorkbook.IsSheetVeryHidden(sheetIx);
        }

        public SheetVisibility GetSheetVisibility(int sheetIx)
        {
            return _wb.GetSheetVisibility(sheetIx);
        }

        [Obsolete]
        public void SetSheetHidden(int sheetIx, SheetVisibility hidden)
        {
            XssfWorkbook.SetSheetHidden(sheetIx, hidden);
        }

        [Obsolete]
        public void SetSheetHidden(int sheetIx, int hidden)
        {
            XssfWorkbook.SetSheetHidden(sheetIx, hidden);
        }

        public void AddToolPack(UDFFinder toopack)
        {
            XssfWorkbook.AddToolPack(toopack);
        }
        public void SetSheetVisibility(int sheetIx, SheetVisibility visibility)
        {
            _wb.SetSheetVisibility(sheetIx, visibility);
        }

        /// <summary>
        /// Returns the spreadsheet version (EXCLE2007) of this workbook
        /// </summary>
        public SpreadsheetVersion SpreadsheetVersion
        {
            get
            {
                return SpreadsheetVersion.EXCEL2007;
            }
        }

        /// <summary>
        /// Gets a bool value that indicates whether the date systems used in the workbook starts in 1904.
        /// The default value is false, meaning that the workbook uses the 1900 date system,
        /// where 1/1/1900 is the first day in the system.
        /// </summary>
        /// <returns>True if the date systems used in the workbook starts in 1904</returns>
        public bool IsDate1904()
        {
            return XssfWorkbook.IsDate1904();
        }

        void IDisposable.Dispose()
        {
            this.Dispose();
        }



        //TODO: missing method isDate1904, isHidden, setHidden

        private sealed class SheetEnumerator<T> : IEnumerator<T> where T : class, ISheet
        {
            private XSSFWorkbook _wb;
            private readonly SXSSFWorkbook _xwb;
            private readonly IEnumerator<ISheet> it;
            public SheetEnumerator(XSSFWorkbook wb, SXSSFWorkbook xwb)
            {
                this._wb = wb;
                this._xwb = xwb;
                //wb.GetEnumerator();
                it = wb.GetEnumerator();
            }

            T IEnumerator<T>.Current
            {
                get
                {
                    XSSFSheet xssfSheet = (XSSFSheet)it.Current;
                    return _xwb.GetSXSSFSheet(xssfSheet) as T;
                }
            }


            object IEnumerator.Current
            {
                get
                {
                    XSSFSheet xssfSheet = (XSSFSheet)it.Current;
                    return _xwb.GetSXSSFSheet(xssfSheet);
                }
            }

            public void Dispose()
            {
                it.Dispose();
            }

            public bool MoveNext()
            {
                return it.MoveNext();
            }

            public void Reset()
            {
                it.Reset();
            }
        }
    }
}
