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
using System.IO;
using System.Linq;
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
        public static readonly int DEFAULT_WINDOW_SIZE = 100;
        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(SXSSFWorkbook));

        public XSSFWorkbook xssfWorkbook;

        private Dictionary<SXSSFSheet, XSSFSheet> _sxFromXHash = new Dictionary<SXSSFSheet, XSSFSheet>();
        private Dictionary<XSSFSheet, SXSSFSheet> _xFromSxHash = new Dictionary<XSSFSheet, SXSSFSheet>();

        private int _randomAccessWindowSize = DEFAULT_WINDOW_SIZE;

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
            get
            {
                throw new NotImplementedException();
            }
        }

        public int FirstVisibleTab
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public int NumberOfSheets
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public short NumberOfFonts
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public short NumCellStyles
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public int NumberOfNames
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public MissingCellPolicy MissingCellPolicy
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public bool IsHidden
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
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

        public SXSSFWorkbook(XSSFWorkbook workbook, int rowAccessWindowSize, bool compressTmpFiles, bool useSharedStringsTable)
        {
            SetRandomAccessWindowSize(rowAccessWindowSize);
            _compressTmpFiles = compressTmpFiles;
            if (workbook == null)
            {
                xssfWorkbook = new XSSFWorkbook();
                _sharedStringSource = useSharedStringsTable ? xssfWorkbook.GetSharedStringSource() : null;
            }
            else
            {
                xssfWorkbook = workbook;
                _sharedStringSource = useSharedStringsTable ? xssfWorkbook.GetSharedStringSource() : null;
                var numberOfSheets = xssfWorkbook.NumberOfSheets;
                for (int i = 0; i < numberOfSheets; i++)
                {
                    XSSFSheet sheet = (XSSFSheet)xssfWorkbook.GetSheetAt(i);
                    CreateAndRegisterSXSSFSheet(sheet);
                }
            }
        }

        #endregion

        private void SetRandomAccessWindowSize(int rowAccessWindowSize)
        {
            if (rowAccessWindowSize == 0 || rowAccessWindowSize < -1)
            {
                throw new ArgumentException("rowAccessWindowSize must be greater than 0 or -1");
            }
            _randomAccessWindowSize = rowAccessWindowSize;
        }

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
            registerSheetMapping(sxSheet, (XSSFSheet)xSheet);
            return sxSheet;
        }

        private void registerSheetMapping(SXSSFSheet sxSheet, XSSFSheet xSheet)
        {
            _sxFromXHash.Add(sxSheet, xSheet);
            _xFromSxHash.Add(xSheet, sxSheet);
        }

        private void DeregisterSheetMapping(XSSFSheet xSheet)
        {
            SXSSFSheet sxSheet = GetSXSSFSheet(xSheet);

            // ensure that the writer is closed in all cases to not have lingering writers
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


        private XSSFSheet getXSSFSheet(SXSSFSheet sheet)
        {
            return _sxFromXHash[sheet];
        }

        private SXSSFSheet GetSXSSFSheet(XSSFSheet sheet)
        {
            return _xFromSxHash[sheet];
        }

        public int GetRandomAccessWindowSize()
        {
            return _randomAccessWindowSize;
        }

        public SheetDataWriter createSheetDataWriter()
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

        protected void InjectData(ZipEntrySource zipEntrySource, Stream outStream)
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
            bool flag = true;
            while ((count = inputStream.Read(chunk, 0, chunk.Length)) >= 0 && flag)
            {
                outputStream.Write(chunk, 0, count);
                if (count == 0) flag = false;
            }
        }

        private static void CopyStreamAndInjectWorksheet(Stream inputStream, Stream outputStream, Stream worksheetData)
        {
            StreamReader inReader = new StreamReader(inputStream, Encoding.UTF8); //TODO: Is it always UTF-8 or do we need to read the xml encoding declaration in the file? If not, we should perhaps use a SAX reader instead.
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
                        //outWriter.Write(s, 0, pos);//this is a little wierd, I think they are just trying to write the '<' character? //TODO this might be my issue
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
            //var inReader = new StreamReader(inputStream, Encoding.UTF8); //TODO: Is it always UTF-8 or do we need to read the xml encoding declaration in the file? If not, we should perhaps use a SAX reader instead.
            //var outWriter = new StreamWriter(outputStream, Encoding.UTF8);
            //bool needsStartTag = true;
            //int c;
            //int pos = 0;
            //string s = "<sheetData";
            //int n = s.Length;
            ////Copy from "in" to "out" up to the string "<sheetData/>" or "</sheetData>" (excluding).
            //while (((c = inReader.Read()) != -1))
            //{
            //    if (c == s[pos])
            //    {
            //        pos++;
            //        if (pos == n)
            //        {
            //            if ("<sheetData".Equals(s))
            //            {
            //                c = inReader.Read();
            //                if (c == -1)
            //                {
            //                    outWriter.Write(s);
            //                    break;
            //                }
            //                if (c == '>')
            //                {
            //                    // Found <sheetData>
            //                    outWriter.Write(s);
            //                    outWriter.Write((char)c);
            //                    s = "</sheetData>";
            //                    n = s.Length;
            //                    pos = 0;
            //                    needsStartTag = false;
            //                    continue;
            //                }
            //                if (c == '/')
            //                {
            //                    // Found <sheetData/
            //                    c = inReader.Read();
            //                    if (c == -1)
            //                    {
            //                        outWriter.Write(s);
            //                        break;
            //                    }
            //                    if (c == '>')
            //                    {
            //                        // Found <sheetData/>
            //                        break;
            //                    }

            //                    outWriter.Write(s);
            //                    outWriter.Write('/');
            //                    outWriter.Write((char)c);
            //                    pos = 0;
            //                    continue;
            //                }

            //                outWriter.Write(s);
            //                outWriter.Write('/');
            //                outWriter.Write((char)c);
            //                pos = 0;
            //                continue;
            //            }
            //            else
            //            {
            //                // Found </sheetData>
            //                break;
            //            }
            //        }
            //    }
            //    else
            //    {
            //        if (pos > 0) outWriter.Write(s, 0, pos);
            //        if (c == s[0])
            //        {
            //            pos = 1;
            //        }
            //        else
            //        {
            //            outWriter.Write((char)c);
            //            pos = 0;
            //        }
            //    }
            //}
            //outWriter.Flush();
            //if (needsStartTag)
            //{
            //    outWriter.Write("<sheetData>\n");
            //    outWriter.Flush();
            //}
            ////Copy the worksheet data to "out".
            //CopyStream(worksheetData, outputStream);
            //outWriter.Write("</sheetData>");
            //outWriter.Flush();
            ////Copy the rest of "in" to "out".
            //while (((c = inReader.Read()) != -1))
            //{
            //    outWriter.Write((char)c);
            //}

            //outWriter.Flush();
        }

        public void SetSheetOrder(string sheetname, int pos)
        {
            xssfWorkbook.SetSheetOrder(sheetname, pos);
        }

        public void SetSelectedTab(int index)
        {
            xssfWorkbook.SetSelectedTab(index);
        }

        public void SetActiveSheet(int sheetIndex)
        {
            throw new NotImplementedException();
        }

        public string GetSheetName(int sheet)
        {
            return xssfWorkbook.GetSheetName(sheet);
        }

        public void SetSheetName(int sheet, string name)
        {
            xssfWorkbook.SetSheetName(sheet, name);
        }

        public int GetSheetIndex(string name)
        {
            return xssfWorkbook.GetSheetIndex(name);
        }

        public int GetSheetIndex(ISheet sheet)
        {
            return xssfWorkbook.GetSheetIndex(getXSSFSheet((SXSSFSheet)sheet));
        }

        public ISheet CreateSheet()
        {
            return CreateAndRegisterSXSSFSheet(xssfWorkbook.CreateSheet());
        }

        public ISheet CreateSheet(string sheetname)
        {
            return CreateAndRegisterSXSSFSheet(xssfWorkbook.CreateSheet(sheetname));
        }

        public ISheet CloneSheet(int sheetNum)
        {
            throw new NotImplementedException();
        }

        public ISheet GetSheetAt(int index)
        {
            return GetSXSSFSheet((XSSFSheet)xssfWorkbook.GetSheetAt(index));
        }

        public ISheet GetSheet(string name)
        {
            return GetSXSSFSheet((XSSFSheet)xssfWorkbook.GetSheet(name));
        }

        public void RemoveSheetAt(int index)
        {
            // Get the sheet to be removed
            var xSheet = (XSSFSheet)xssfWorkbook.GetSheetAt(index);
            var sxSheet = GetSXSSFSheet(xSheet);

            // De-register it
            xssfWorkbook.RemoveSheetAt(index);
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
            return xssfWorkbook.CreateFont();
        }

        [Obsolete("deprecated in poi 3.16")]
        public IFont FindFont(short boldWeight, short color, short fontHeight, string name, bool italic, bool strikeout, FontSuperScript typeOffset, FontUnderlineType underline)
        {
            return xssfWorkbook.FindFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
        }

        public IFont GetFontAt(short idx)
        {
            return xssfWorkbook.GetFontAt(idx);
        }

        public ICellStyle CreateCellStyle()
        {
            return xssfWorkbook.CreateCellStyle();
        }

        public ICellStyle GetCellStyleAt(short idx)
        {
            return xssfWorkbook.GetCellStyleAt(idx);
        }

        public void Write(Stream stream)
        {
            FlushSheets();

            //Save the template
            var tmplFile = TempFile.CreateTempFile("poi-sxssf-template", ".xlsx");
            try
            {
                //TODO: may just want to be a filestream
                var os = new FileStream(tmplFile.FullName, FileMode.Open, FileAccess.Write);
                try
                {
                    xssfWorkbook.Write(os);
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
        protected void FlushSheets()
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
            return xssfWorkbook.GetName(name);
        }
        [Obsolete("Deprecated in 3.16 throws an error.")]
        public IName GetNameAt(int nameIndex)
        {
            return xssfWorkbook.GetNameAt(nameIndex);
        }

        public IName CreateName()
        {
            return xssfWorkbook.CreateName();
        }

        [Obsolete("deprecated in 3.16 New projects should avoid accessing named ranges by index. GetName(String)} instead.")]
        public int GetNameIndex(string name)
        {
            return xssfWorkbook.GetNameIndex(name);
        }

        [Obsolete("deprecated in 3.16 New projects should use RemoveName(Name)")]
        public void RemoveName(int index)
        {
            xssfWorkbook.RemoveName(index);
        }

        [Obsolete("deprecated in 3.16 New projects should use RemoveName(IName Name)")]
        public void RemoveName(string name)
        {
            xssfWorkbook.RemoveName(name);
        }

        public int LinkExternalWorkbook(string name, IWorkbook workbook)
        {
            throw new NotImplementedException();
        }

        public void SetPrintArea(int sheetIndex, string reference)
        {
            xssfWorkbook.SetPrintArea(sheetIndex, reference);
        }

        public void SetPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow)
        {
            xssfWorkbook.SetPrintArea(sheetIndex, startColumn, endColumn, startRow, endRow);
        }

        public string GetPrintArea(int sheetIndex)
        {
            return xssfWorkbook.GetPrintArea(sheetIndex);
        }

        public void RemovePrintArea(int sheetIndex)
        {
            xssfWorkbook.RemovePrintArea(sheetIndex);
        }

        //TODO: missing methods Get/Set MissingCellPolicy


        public IDataFormat CreateDataFormat()
        {
            return xssfWorkbook.CreateDataFormat();
        }

        public int AddPicture(byte[] pictureData, PictureType format)
        {
            return xssfWorkbook.AddPicture(pictureData, format);
        }

        public IList GetAllPictures()
        {
            return xssfWorkbook.GetAllPictures();
        }

        public ICreationHelper GetCreationHelper()
        {
            return new SXSSFCreationHelper(this);
        }

        //TODO: missing method isDate1904, isHidden, setHidden

        public bool IsSheetHidden(int sheetIx)
        {
            return xssfWorkbook.IsSheetHidden(sheetIx);
        }

        public bool IsSheetVeryHidden(int sheetIx)
        {
            return xssfWorkbook.IsSheetVeryHidden(sheetIx);
        }

        public void SetSheetHidden(int sheetIx, SheetState hidden)
        {
            xssfWorkbook.SetSheetHidden(sheetIx, hidden);
        }

        public void SetSheetHidden(int sheetIx, int hidden)
        {
            xssfWorkbook.SetSheetHidden(sheetIx, hidden);
        }

        public void AddToolPack(UDFFinder toopack)
        {
            xssfWorkbook.AddToolPack(toopack);
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
            xssfWorkbook.Close();
        }

        //TODO: missing methods setForceFormulaRecalculation, GetForceFormulaRecalulation, GetSpreadsheetVersion
    }
}
