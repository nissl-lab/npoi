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
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.OpenXml4Net.Util;
using NPOI.SS.Formula.Udf;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    /// <summary>
    /// Streaming version of the XSSFWorkbook, originally implemented in the "BigGridDemo".
    /// </summary>
    public class SXSSFWorkbook : IWorkbook
    {
        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(SXSSFWorkbook));

        public const int DEFAULT_WINDOW_SIZE = 100;
        
        public XSSFWorkbook XssfWorkbook;

        private Dictionary<SXSSFSheet, XSSFSheet> _sxFromXHash = new Dictionary<SXSSFSheet, XSSFSheet>();
        private Dictionary<XSSFSheet, SXSSFSheet> _xFromSxHash = new Dictionary<XSSFSheet, SXSSFSheet>();

        private int _randomAccessWindowSize = DEFAULT_WINDOW_SIZE;

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
        /// shared string table - a cache of strings in this workbook.
        /// </summary>
        private SharedStringsTable _sharedStringSource;

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

        public short NumCellStyles
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


        #region Constructors

        public SXSSFWorkbook() : this(null)
        {

        }

        public SXSSFWorkbook(XSSFWorkbook workbook) : this(workbook, DEFAULT_WINDOW_SIZE)
        {

        }

        public SXSSFWorkbook(XSSFWorkbook workbook, int rowAccessWindowSize)
            : this(workbook, rowAccessWindowSize, false)
        {

        }

        public SXSSFWorkbook(XSSFWorkbook workbook, int rowAccessWindowSize, bool compressTmpFiles)
            : this(workbook, rowAccessWindowSize, compressTmpFiles, false)
        {

        }

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
                XssfWorkbook = new XSSFWorkbook();
                _sharedStringSource = useSharedStringsTable ? XssfWorkbook.GetSharedStringSource() : null;
            }
            else
            {
                XssfWorkbook = workbook;
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
                sxSheet._writer.Close();
            }
            catch (IOException e)
            {
                // ignore exception here
            }

            _sxFromXHash.Remove(sxSheet);

            _xFromSxHash.Remove(xSheet);
        }


        private XSSFSheet GetXSSFSheet(SXSSFSheet sheet)
        {
            return _sxFromXHash[sheet];
        }

        private SXSSFSheet GetSXSSFSheet(XSSFSheet sheet)
        {
            return _xFromSxHash[sheet];
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

        private void InjectData(ZipEntrySource zipEntrySource, Stream outStream)
        {
            try
            {
                ZipOutputStream zos = new ZipOutputStream(outStream);
                try
                {
                    var en = zipEntrySource.Entries;
                    while (en.MoveNext())
                    {
                        var ze = (ZipEntry)en.Current;
                        zos.PutNextEntry(new ZipEntry(ze.Name));
                        var inputStream = zipEntrySource.GetInputStream(ze);
                        XSSFSheet xSheet = GetSheetFromZipEntryName(ze.Name);
                        if (xSheet != null)
                        {
                            SXSSFSheet sxSheet = GetSXSSFSheet(xSheet);
                            var xis = sxSheet.getWorksheetXMLInputStream();
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
                            CopyStream(inputStream, zos);
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
                zipEntrySource.Close();
            }
        }

        private static void CopyStream(Stream inputStream, Stream outputStream)
        {
            var chunk = new byte[1024];
            int count;

            while ((count = inputStream.Read(chunk, 0, chunk.Length)) > 0)
            {
                outputStream.Write(chunk, 0, count);
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
            CopyStream(worksheetData, outputStream);
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
            return CreateAndRegisterSXSSFSheet(XssfWorkbook.CreateSheet(sheetname));
        }

        public ISheet CloneSheet(int sheetNum)
        {
            throw new NotImplementedException();
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
                sxSheet.dispose();
            }
            catch (IOException e)
            {
                logger.Log(POILogger.WARN, e);
            }
        }

        public IEnumerator GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public void SetRepeatingRowsAndColumns(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow)
        {
            throw new NotImplementedException();
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

        public IFont GetFontAt(short idx)
        {
            return XssfWorkbook.GetFontAt(idx);
        }

        public ICellStyle CreateCellStyle()
        {
            return XssfWorkbook.CreateCellStyle();
        }

        public ICellStyle GetCellStyleAt(short idx)
        {
            return XssfWorkbook.GetCellStyleAt(idx);
        }

        public void Write(Stream stream)
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
                ZipEntrySource source = new ZipFileZipEntrySource(new ZipFile(tmplFile.FullName));
                InjectData(source, stream);
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
                sheet.flushRows();
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
                    success = sheet.dispose() && success;
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
        [Obsolete("Deprecated in 3.16 throws an error.")]
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

        public void SetSheetHidden(int sheetIx, SheetState hidden)
        {
            XssfWorkbook.SetSheetHidden(sheetIx, hidden);
        }

        public void SetSheetHidden(int sheetIx, int hidden)
        {
            XssfWorkbook.SetSheetHidden(sheetIx, hidden);
        }

        public void AddToolPack(UDFFinder toopack)
        {
            XssfWorkbook.AddToolPack(toopack);
        }

        public void Close()
        {
            // ensure that any lingering writer is closed
            foreach (SXSSFSheet sheet in _xFromSxHash.Values)
            {
                try
                {
                    sheet._writer.Close();
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

        //TODO: missing methods from POI 3.16 setForceFormulaRecalculation, GetForceFormulaRecalulation, GetSpreadsheetVersion
        //TODO: missing method isDate1904, isHidden, setHidden
    }
}
