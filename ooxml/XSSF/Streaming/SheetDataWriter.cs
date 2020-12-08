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
using System.Globalization;
using System.IO;
using System.Text;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SheetDataWriter
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(SheetDataWriter));

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
        private SharedStringsTable _sharedStringSource;
        private StreamWriter _outputWriter;

        public SheetDataWriter()
        {
            TemporaryFileInfo = CreateTempFile();
            OutputStream = CreateWriter(TemporaryFileInfo);
            _outputWriter = new StreamWriter(OutputStream, Encoding.UTF8);
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
            Stream outputStream = null;
            try
            {
                outputStream = DecorateOutputStream(fos);
            }
            catch (Exception)
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
                _outputWriter.Flush();
                OutputStream.Flush();
            }
            catch (Exception)
            {

            }
            try
            {
                OutputStream.Close();
            }
            catch (Exception)
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
            Stream fis = new FileStream(TemporaryFileInfo.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
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

            _outputWriter.Flush();
            OutputStream.Flush();
        }

        private void BeginRow(int rownum, SXSSFRow row)
        {
            WriteAsBytes("<row r=\"");
            WriteAsBytes(rownum + 1);
            WriteAsBytes("\"");

            if (row.HasCustomHeight())
            {
                WriteAsBytes(" customHeight=\"true\"  ht=\"");
                WriteAsBytes(row.HeightInPoints);
                WriteAsBytes("\"");
            }
            if (row.ZeroHeight)
            {
                WriteAsBytes(" hidden=\"true\"");
            }
            if (row.IsFormatted)
            {
                WriteAsBytes(" s=\"");
                WriteAsBytes(row.RowStyle.Index);
                WriteAsBytes("\"");
                WriteAsBytes(" customFormat=\"1\"");
            }

            if (row.OutlineLevel != 0)
            {
                WriteAsBytes(" outlineLevel=\"");
                WriteAsBytes(row.OutlineLevel);
                WriteAsBytes("\"");
            }
            if (row.Hidden != null)
            {
                WriteAsBytes(" hidden=\"");
                WriteAsBytes(row.Hidden.Value ? "1" : "0");
                WriteAsBytes("\"");
            }
            if (row.Collapsed != null)
            {
                WriteAsBytes(" collapsed=\"");
                WriteAsBytes(row.Collapsed.Value ? "1" : "0");
                WriteAsBytes("\"");
            }

            WriteAsBytes(">\n");

            RowNum = rownum;
        }

        private void EndRow()
        {
            WriteAsBytes("</row>\n");
        }

        public void WriteCell(int columnIndex, ICell cell)
        {
            if (cell == null)
            {
                return;
            }
            var cellRef = new CellReference(RowNum, columnIndex).FormatAsString();
            WriteAsBytes("<c r=\"");
            WriteAsBytes(cellRef);
            WriteAsBytes("\"");

            if (cell.CellStyle.Index != 0)
            {
                // need to convert the short to unsigned short as the indexes can be up to 64k
                // ideally we would use int for this index, but that would need changes to some more 
                // APIs
                WriteAsBytes(" s=\"");
                WriteAsBytes(cell.CellStyle.Index & 0xffff);
                WriteAsBytes("\"");
            }
            switch (cell.CellType)
            {
                case CellType.Blank:
                    {
                        WriteAsBytes(">");
                        break;
                    }
                case CellType.Formula:
                    {
                        WriteAsBytes(">");
                        WriteAsBytes("<f>");

                        OutputQuotedString(cell.CellFormula);

                        WriteAsBytes("</f>");

                        switch (cell.GetCachedFormulaResultTypeEnum())
                        {
                            case CellType.Numeric:
                                double nval = cell.NumericCellValue;
                                if (!Double.IsNaN(nval))
                                {
                                    WriteAsBytes("<v>");
                                    WriteAsBytes(nval);
                                    WriteAsBytes("</v>");
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

                            WriteAsBytes(" t=\"");
                            WriteAsBytes("s");
                            WriteAsBytes("\">");
                            WriteAsBytes("<v>");
                            WriteAsBytes(sRef);
                            WriteAsBytes("</v>");
                        }
                        else
                        {
                            WriteAsBytes(" t=\"inlineStr\">");
                            WriteAsBytes("<is><t");

                            if (HasLeadingTrailingSpaces(cell.StringCellValue))
                            {
                                WriteAsBytes(" xml:space=\"preserve\"");
                            }

                            WriteAsBytes(">");

                            OutputQuotedString(cell.StringCellValue);

                            WriteAsBytes("</t></is>");
                        }
                        break;
                    }
                case CellType.Numeric:
                    {
                        WriteAsBytes(" t=\"n\">");
                        WriteAsBytes("<v>");
                        WriteAsBytes(cell.NumericCellValue);
                        WriteAsBytes("</v>");
                        break;
                    }
                case CellType.Boolean:
                    {
                        WriteAsBytes(" t=\"b\">");
                        WriteAsBytes("<v>");
                        WriteAsBytes(cell.BooleanCellValue ? "1" : "0");
                        WriteAsBytes("</v>");
                        break;
                    }
                case CellType.Error:
                    {
                        FormulaError error = FormulaError.ForInt(cell.ErrorCellValue);

                        WriteAsBytes(" t=\"e\">");
                        WriteAsBytes("<v>");
                        WriteAsBytes(error.String);
                        WriteAsBytes("</v>");
                        break;
                    }
                default:
                    {
                        throw new InvalidOperationException("Invalid cell type: " + cell.CellType);
                    }
            }
            WriteAsBytes("</c>");
        }

        private void WriteAsBytes(string text)
        {
            _outputWriter.Write(text);
        }

        private void WriteAsBytes(ArraySegment<char> chars)
        {
            _outputWriter.Write(chars.Array, chars.Offset, chars.Count);
        }

        private void WriteAsBytes(int value)
        {
            _outputWriter.Write(value);
        }

        private void WriteAsBytes(float value)
        {
            _outputWriter.Write(value.ToString(CultureInfo.InvariantCulture));
        }

        private void WriteAsBytes(double value)
        {
            _outputWriter.Write(value.ToString(CultureInfo.InvariantCulture));
        }

        /**
         * @return  whether the string has leading / trailing spaces that
         *  need to be preserved with the xml:space=\"preserve\" attribute
         */
        private bool HasLeadingTrailingSpaces(string str)
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
        protected void OutputQuotedString(string s)
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
                        if (counter > last)
                        {
                            WriteAsBytes(GetSubArray(chars, last, counter - last));
                        }
                        last = counter + 1;

                        WriteAsBytes("&lt;");
                        break;
                    case '>':
                        if (counter > last)
                        {
                            WriteAsBytes(GetSubArray(chars, last, counter - last));
                        }
                        last = counter + 1;
                        WriteAsBytes("&gt;");
                        break;
                    case '&':
                        if (counter > last)
                        {
                            WriteAsBytes(GetSubArray(chars, last, counter - last));
                        }
                        last = counter + 1;
                        WriteAsBytes("&amp;");
                        break;
                    case '"':
                        if (counter > last)
                        {
                            WriteAsBytes(GetSubArray(chars, last, counter - last));
                        }
                        last = counter + 1;
                        WriteAsBytes("&quot;");
                        break;
                    // Special characters
                    case '\n':
                    case '\r':
                        if (counter > last)
                        {
                            WriteAsBytes(GetSubArray(chars, last, counter - last));
                        }
                        WriteAsBytes("&#xa;");
                        last = counter + 1;
                        break;
                    case '\t':
                        if (counter > last)
                        {
                            WriteAsBytes(GetSubArray(chars, last, counter - last));
                        }
                        WriteAsBytes("&#x9;");
                        last = counter + 1;
                        break;
                    case (char)0xa0:
                        if (counter > last)
                        {
                            WriteAsBytes(GetSubArray(chars, last, counter - last));
                        }
                        WriteAsBytes("&#xa0;");
                        last = counter + 1;
                        break;
                    default:
                        // YK: XmlBeans silently replaces all ISO control characters ( < 32) with question marks.
                        // the same rule applies to unicode surrogates and "not a character" symbols.
                        if (c < ' ' || Char.IsLowSurrogate(c) || Char.IsHighSurrogate(c) || '\uFFFE' <= c)
                        {
                            if (counter > last)
                            {
                                WriteAsBytes(GetSubArray(chars, last, counter - last));
                            }
                            WriteAsBytes("?");
                            last = counter + 1;
                        }
                        else if (c > 127)
                        {
                            if (counter > last)
                            {
                                WriteAsBytes(GetSubArray(chars, last, counter - last));
                            }
                            last = counter + 1;
                            // If the character is outside of UTF8, write the
                            // numeric value.
                            WriteAsBytes("&#");
                            WriteAsBytes(c);
                            WriteAsBytes(";");
                        }
                        break;
                }
            }
            if (last < length)
            {
                WriteAsBytes(GetSubArray(chars, last, length - last));
            }
        }

        private static ArraySegment<char> GetSubArray(char[] oldArray, int skip, int take)
        {
            return new ArraySegment<char>(oldArray, skip, take);
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
                OutputStream.Close();
            }
            finally
            {
                TemporaryFileInfo.Delete();
                ret = !File.Exists(TemporaryFileInfo.FullName);
                TemporaryFileInfo.Refresh();
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
