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
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace NPOI.XSSF.Streaming
{
    public class SheetDataWriter : ICloseable
    {
        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(SheetDataWriter));

        protected FileInfo TemporaryFileInfo { get; set; }
        protected Stream OutputStream { get; private set; }
        private int RowNum { get; set; }
        public int NumberOfFlushedRows { get; set; }
        public int LowestIndexOfFlushedRows { get; set; } = -1; // meaningful only of _numberOfFlushedRows>0
        public int NumberOfCellsOfLastFlushedRow { get; set; } // meaningful only of _numberOfFlushedRows>0
        public int NumberLastFlushedRow = -1; // meaningful only of _numberOfFlushedRows>0

        /**
         * Table of strings shared across this workbook.
         * If two cells contain the same string, then the cell value is the same index into SharedStringsTable
         */
        private readonly SharedStringsTable _sharedStringSource;
        private readonly StreamWriter _out;

        public SheetDataWriter()
        {
            TemporaryFileInfo = CreateTempFile();
            OutputStream = CreateWriter(TemporaryFileInfo);
            _out = new StreamWriter(OutputStream);
        }
        public SheetDataWriter(SharedStringsTable sharedStringsTable) : this()
        {
            _sharedStringSource = sharedStringsTable;
        }
        /**
         * Create a temp file to write sheet data. 
         * By default, temp files are created in the default temporary-file directory
         * with a prefix "poi-sxssf-sheet" and suffix ".xml".  Subclasses can override 
         * it and specify a different temp directory or filename or suffix, e.g. <code>.gz</code>
         * 
         * @return temp file to write sheet data
         */

        public virtual FileInfo CreateTempFile()
        {
            return TempFile.CreateTempFile("poi-sxssf-sheet", ".xml");
        }

        /**
         * Create a writer for the sheet data.
         * 
         * @param  fd the file to write to
         */
        public virtual Stream CreateWriter(FileInfo fd)
        {

            FileStream fos = new FileStream(fd.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            Stream outputStream;
            try
            {
                outputStream = DecorateOutputStream(fos);
            }
            catch
            {
                fos.Close();
                throw;
            }

            return outputStream;
        }

        /**
         * Override this to translate (such as encrypt or compress) the file output stream
         * as it is being written to disk.
         * The default behavior is to to pass the stream through unmodified.
         *
         * @param fos  the stream to decorate
         * @return a decorated stream
         * @throws IOException
         * @see #decorateInputStream(FileInputStream)
         */
        protected virtual Stream DecorateOutputStream(Stream fos)
        {
            return fos;
        }

        /**
         * flush and close the temp data writer. 
         * This method <em>must</em> be invoked before calling {@link #getWorksheetXMLInputStream()}
         */
        public void Close()
        {
            try
            {
                _out.Dispose();
                OutputStream.Dispose();
            }
            catch
            {

            }
        }

        public FileInfo TempFileInfo
        {
            get
            {
                return TemporaryFileInfo;
            }
        }
        /**
         * @return a stream to read temp file with the sheet data
         */
        public Stream GetWorksheetXmlInputStream()
        {
            FileStream fis = new FileStream(TemporaryFileInfo.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            try
            {
                return DecorateInputStream(fis);
            }
            catch (IOException)
            {
                fis.Close();
                throw;
            }
        }

        /**
         * Override this to translate (such as decrypt or expand) the file input stream
         * as it is being read from disk.
         * The default behavior is to to pass the stream through unmodified.
         *
         * @param fis  the stream to decorate
         * @return a decorated stream
         * @throws IOException
         * @see #decorateOutputStream(FileOutputStream)
         */
        protected virtual Stream DecorateInputStream(Stream fis)
        {
            return fis;
        }

        protected void FinalizeWriter()
        {
            TemporaryFileInfo.Delete();
            if (File.Exists(TemporaryFileInfo.FullName))
            {
                logger.Log(POILogger.ERROR, "Can't delete temporary encryption file: " + TemporaryFileInfo);
            }
        }

        /**
         * Write a row to the file
         *
         * @param rownum 0-based row number
         * @param row    a row
         */
        public void WriteRow(int rownum, SXSSFRow row)
        {
            BeginRow(rownum, row);
            
            using (var cells = row.AllCellsIterator())
            {
                int columnIndex = 0;
                while (cells.MoveNext())
                {
                    WriteCell(columnIndex++, cells.Current);
                }
                EndRow();
            }
            
            if (LowestIndexOfFlushedRows == -1 || LowestIndexOfFlushedRows > rownum)
            {
                LowestIndexOfFlushedRows = rownum;
                NumberOfFlushedRows++;
            }
        }

        public void FlushRows(int rowCount, int lastRowNum, int lastRowCellsCount)
        {

            NumberLastFlushedRow = Math.Max(lastRowNum, NumberLastFlushedRow);
            NumberOfCellsOfLastFlushedRow = lastRowCellsCount;

            _out.Flush();
            OutputStream.Flush();
        }

        private void BeginRow(int rownum, SXSSFRow row)
        {
            _out.Write("<row");
            WriteAttribute("r", (rownum + 1).ToString());

            if (row.HasCustomHeight())
            {
                WriteAttribute("customHeight", "true");
                WriteAttribute("ht", row.HeightInPoints.ToString(CultureInfo.InvariantCulture));
            }
            if (row.ZeroHeight)
            {
                WriteAttribute("hidden", "true");
            }
            if (row.IsFormatted)
            {
                WriteAttribute("s", row.RowStyleIndex.ToString());
                WriteAttribute("customFormat", "1");
            }

            if (row.OutlineLevel != 0)
            {
                WriteAttribute("outlineLevel", row.OutlineLevel.ToString());
            }
            if (row.Hidden != null)
            {
                WriteAttribute("hidden", (row.Hidden?? true) ? "1" : "0");
            }
            if (row.Collapsed != null)
            {
                WriteAttribute("collapsed", (row.Collapsed ?? true) ? "1" : "0");
            }

            _out.Write(">\n");

            RowNum = rownum;
        }

        private void EndRow()
        {
            _out.Write("</row>\n");
        }

        public void WriteCell(int columnIndex, ICell cell)
        {
            if (cell == null)
            {
                return;
            }
            var cellRef = new CellReference(RowNum, columnIndex).FormatAsString();
            _out.Write("<c");
            WriteAttribute("r", cellRef);
            ICellStyle cellStyle = cell.CellStyle;
            if (cellStyle.Index != 0)
            {
                // need to convert the short to unsigned short as the indexes can be up to 64k
                // ideally we would use int for this index, but that would need changes to some more 
                // APIs
                WriteAttribute("s", (cellStyle.Index & 0xffff).ToString());
            }
            switch (cell.CellType)
            {
                case CellType.Blank:
                    {
                        _out.Write('>');
                        break;
                    }
                case CellType.Formula:
                    {
                        _out.Write("><f>");
                        OutputQuotedString(cell.CellFormula);
                        _out.Write("</f>");

                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.Numeric:
                                double nval = cell.NumericCellValue;
                                if (!Double.IsNaN(nval))
                                {
                                    _out.Write("<v>");
                                    _out.Write(nval.ToString(CultureInfo.InvariantCulture));
                                    _out.Write("</v>");
                                }
                                break;
                            default:
                                break;
                        }
                        break;
                    }
                case CellType.String:
                    {
                        if (_sharedStringSource != null)
                        {
                            XSSFRichTextString rt = new XSSFRichTextString(cell.StringCellValue);
                            int sRef = _sharedStringSource.AddEntry(rt.GetCTRst());

                            WriteAttribute("t", ST_CellType.s.ToString());
                            _out.Write("><v>");
                            _out.Write(sRef.ToString());
                            _out.Write("</v>");
                        }
                        else
                        {
                            WriteAttribute("t", "inlineStr");
                            _out.Write("><is><t");

                            if (HasLeadingTrailingSpaces(cell.StringCellValue))
                            {
                                WriteAttribute("xml:space", "preserve");
                            }

                            _out.Write(">");

                            OutputQuotedString(cell.StringCellValue);

                            _out.Write("</t></is>");
                        }
                        break;
                    }
                case CellType.Numeric:
                    {
                        WriteAttribute("t", "n");
                        _out.Write("><v>");
                        _out.Write(cell.NumericCellValue.ToString(CultureInfo.InvariantCulture));
                        _out.Write("</v>");
                        break;
                    }
                case CellType.Boolean:
                    {
                        WriteAttribute("t", "b");
                        _out.Write("><v>");
                        _out.Write(cell.BooleanCellValue ? "1" : "0");
                        _out.Write("</v>");
                        break;
                    }
                case CellType.Error:
                    {
                        FormulaError error = FormulaError.ForInt(cell.ErrorCellValue);

                        WriteAttribute("t", "e");
                        _out.Write("><v>");
                        _out.Write(error.String);
                        _out.Write("</v>");
                        break;
                    }
                default:
                    {
                        throw new InvalidOperationException("Invalid cell type: " + cell.CellType);
                    }
            }
            _out.Write("</c>");
        }

        private void WriteAttribute(string name, string value)
        {
            _out.Write(' ');
            _out.Write(name);
            _out.Write("=\"");
            _out.Write(value);
            _out.Write('\"');
        }

        /**
         * @return  whether the string has leading / trailing spaces that
         *  need to be preserved with the xml:space=\"preserve\" attribute
         */
        private static bool HasLeadingTrailingSpaces(string str)
        {
            if (!string.IsNullOrEmpty(str))
            {
                char firstChar = str[0];
                char lastChar = str[str.Length - 1];
                return Character.isWhitespace(firstChar) || Character.isWhitespace(lastChar);
            }
            return false;
        }

        //Taken from jdk1.3/src/javax/swing/text/html/HTMLWriter.java
        internal void OutputQuotedString(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return;
            }

            char[] chars = s.ToCharArray();
            int last = 0;
            int length = s.Length;
            for (int counter = 0; counter < length; counter++)
            {
                char c = chars[counter];
                switch (c)
                {
                    case '<':
                        WriteLastChars(_out, chars, last, counter);
                        last = counter + 1;

                        _out.Write("&lt;");
                        break;
                    case '>':
                        WriteLastChars(_out, chars, last, counter);
                        last = counter + 1;
                        _out.Write("&gt;");
                        break;
                    case '&':
                        WriteLastChars(_out, chars, last, counter);
                        last = counter + 1;
                        _out.Write("&amp;");
                        break;
                    case '"':
                        WriteLastChars(_out, chars, last, counter);
                        last = counter + 1;
                        _out.Write("&quot;");
                        break;
                    // Special characters
                    case '\n':
                        WriteLastChars(_out, chars, last, counter);
                        _out.Write("&#xa;");
                        last = counter + 1;
                        break;
                    case '\r':
                        WriteLastChars(_out, chars, last, counter);
                        _out.Write("&#xd;");
                        last = counter + 1;
                        break;
                    case '\t':
                        WriteLastChars(_out, chars, last, counter);
                        _out.Write("&#x9;");
                        last = counter + 1;
                        break;
                    case (char)0xa0:
                        WriteLastChars(_out, chars, last, counter);
                        _out.Write("&#xa0;");
                        last = counter + 1;
                        break;
                    default:
                        // YK: XmlBeans silently replaces all ISO control characters ( < 32) with question marks.
                        // the same rule applies to unicode surrogates and "not a character" symbols.
                        if (ReplaceWithQuestionMark(c)) 
                        {
                            WriteLastChars(_out, chars, last, counter);
                            _out.Write('?');
                            last = counter + 1;
                        }
                        else if (Char.IsHighSurrogate(c) || Char.IsLowSurrogate(c))
                        {
                            WriteLastChars(_out, chars, last, counter);
                            _out.Write(c);
                            last = counter + 1;
                        }
                        else if (c > 127)
                        {
                            WriteLastChars(_out, chars, last, counter);
                            last = counter + 1;
                            // If the character is outside of ascii, write the
                            // numeric value.
                            _out.Write("&#");
                            _out.Write(((int) c).ToString());
                            _out.Write(";");
                        }
                        break;
                }
            }
            if (last < length)
            {
                _out.Write(chars, last, length - last);
            }
        }
        internal static bool ReplaceWithQuestionMark(char c) {
            return c < ' ' || ('\uFFFE' <= c && c <= '\uFFFF');
        }
        private static void WriteLastChars(StreamWriter out1, char[] chars, int last, int counter)
        {
            if (counter > last) {
                out1.Write(chars, last, counter - last);
            }
        }

        /**
         * Deletes the temporary file that backed this sheet on disk.
         * @return true if the file was deleted, false if it wasn't.
         */
        public bool Dispose()
        {
            bool ret;
            try
            {
                _out.Close();
                OutputStream.Close();
            }
            finally
            {
                try
                {
                    TemporaryFileInfo.Delete();
                    ret = !File.Exists(TemporaryFileInfo.FullName);
                    TemporaryFileInfo.Refresh();
                }
                catch(Exception)
                {
                    ret = false;
                }
            }
            return ret;
        }

        public string TemporaryFilePath()
        {
            if (TemporaryFileInfo != null)
            {
                return TemporaryFileInfo.FullName;
            }
            else
            {
                return string.Empty;
            }
        }
    }
}
