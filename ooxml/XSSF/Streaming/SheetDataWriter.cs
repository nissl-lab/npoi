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
using System.IO;
using System.Linq;
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
        protected Stream OutputStream { get; set; }
        private int RowNum { get; set; }
        public int NumberOfFlushedRows { get; set; }
        public int LowestIndexOfFlushedRows { get; set; } // meaningful only of _numberOfFlushedRows>0
        public int NumberOfCellsOfLastFlushedRow { get; set; } // meaningful only of _numberOfFlushedRows>0
        public int NumberLastFlushedRow = -1; // meaningful only of _numberOfFlushedRows>0

        /**
 * Table of strings shared across this workbook.
 * If two cells contain the same string, then the cell value is the same index into SharedStringsTable
 */
        private SharedStringsTable _sharedStringSource;

        public SheetDataWriter()
        {
            TemporaryFileInfo = CreateTempFile();
            OutputStream = CreateWriter(TemporaryFileInfo);
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
            catch (Exception e)
            {
                fos.Close();
                throw e;
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
                OutputStream.Flush();
            }
            catch (Exception e)
            {

            }
            try
            {
                OutputStream.Close();
            }
            catch (Exception e)
            {

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
            catch (IOException e)
            {
                fis.Close();
                throw e;
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



        protected void Finalize()
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
        public void WriteRow(int rownum, IRow row)
        {
            if (NumberOfFlushedRows == 0)
            {
                LowestIndexOfFlushedRows = rownum;
            }
                
            NumberLastFlushedRow = Math.Max(rownum, NumberLastFlushedRow);
            NumberOfCellsOfLastFlushedRow = row.LastCellNum;
            NumberOfFlushedRows++;
            BeginRow(rownum, row);
            var cells = row.GetEnumerator();
            int columnIndex = 0;
            while (cells.MoveNext())
            {
                WriteCell(columnIndex++, cells.Current);
            }
            EndRow();
        }

        private void BeginRow(int rownum, IRow row)
        {
            WriteAsBytes(OutputStream, "<row r=\"" + (rownum + 1) + "\"");

            if (row.HasCustomHeight())
            {
                WriteAsBytes(OutputStream, " customHeight=\"true\"  ht=\"" + row.HeightInPoints + "\"");

            }
            if (row.ZeroHeight)
            {
                WriteAsBytes(OutputStream, " hidden=\"true\"");

            }
            if (row.IsFormatted)
            {
                WriteAsBytes(OutputStream, " s=\"" + row.RowStyle.Index + "\"");

                WriteAsBytes(OutputStream, " customFormat=\"1\"");

            }

            if (row.OutlineLevel != 0)
            {
                WriteAsBytes(OutputStream, " outlineLevel=\"" + row.OutlineLevel + "\"");

            }
            if (row.Hidden != null)
            {
                WriteAsBytes(OutputStream, " hidden=\"" + (row.Hidden.Value ? "1" : "0") + "\"");

            }
            if (row.Collapsed != null)
            {
                WriteAsBytes(OutputStream, " collapsed=\"" + (row.Collapsed.Value ? "1" : "0") + "\"");

            }

            WriteAsBytes(OutputStream, ">\n");

            this.RowNum = rownum;
        }

        private void EndRow()
        {
            WriteAsBytes(OutputStream, "</row>\n");

        }

        public void WriteCell(int columnIndex, ICell cell)
        {
            if (cell == null)
            {
                return;
            }
            string cellRef = new CellReference(RowNum, columnIndex).FormatAsString();
            WriteAsBytes(OutputStream, "<c r=\"" + cellRef + "\"");

            if (cell.CellStyle.Index != 0)
            {
                // need to convert the short to unsigned short as the indexes can be up to 64k
                // ideally we would use int for this index, but that would need changes to some more 
                // APIs
                WriteAsBytes(OutputStream, " s=\"" + (cell.CellStyle.Index & 0xffff) + "\"");
            }
            switch (cell.CellType)
            {
                case CellType.Blank:
                    {
                        WriteAsBytes(OutputStream, ">");
                        break;
                    }
                case CellType.Formula:
                    {
                        WriteAsBytes(OutputStream, ">");
                        WriteAsBytes(OutputStream, "<f>");

                        OutputQuotedString(cell.CellFormula);

                        WriteAsBytes(OutputStream, "</f>");

                        switch (cell.GetCachedFormulaResultTypeEnum())
                        {
                            case CellType.Numeric:
                                double nval = cell.NumericCellValue;
                                if (!Double.IsNaN(nval))
                                {
                                    WriteAsBytes(OutputStream, "<v>" + nval + "</v>");
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

                            WriteAsBytes(OutputStream, " t=\"" + ST_CellType.s + "\">");
                            WriteAsBytes(OutputStream, "<v>");
                            WriteAsBytes(OutputStream, sRef.ToString());
                            WriteAsBytes(OutputStream, "</v>");
                        }
                        else
                        {
                            WriteAsBytes(OutputStream, " t=\"inlineStr\">");
                            WriteAsBytes(OutputStream, "<is><t");

                            if (HasLeadingTrailingSpaces(cell.StringCellValue))
                            {
                                WriteAsBytes(OutputStream, " xml:space=\"preserve\"");
                            }

                            WriteAsBytes(OutputStream, ">");

                            OutputQuotedString(cell.StringCellValue);

                            WriteAsBytes(OutputStream, "</t></is>");
                        }
                        break;
                    }
                case CellType.Numeric:
                    {
                        WriteAsBytes(OutputStream, " t=\"n\">");
                        WriteAsBytes(OutputStream, "<v>" + cell.NumericCellValue + "</v>");
                        break;
                    }
                case CellType.Boolean:
                    {
                        WriteAsBytes(OutputStream, " t=\"b\">");
                        WriteAsBytes(OutputStream, "<v>" + (cell.BooleanCellValue ? "1" : "0") + "</v>");
                        break;
                    }
                case CellType.Error:
                    {
                        FormulaError error = FormulaError.ForInt(cell.ErrorCellValue);

                        WriteAsBytes(OutputStream, " t=\"e\">");
                        WriteAsBytes(OutputStream, "<v>" + error.String + "</v>");
                        break;
                    }
                default:
                    {
                        throw new InvalidOperationException("Invalid cell type: " + cell.CellType);
                    }
            }
            WriteAsBytes(OutputStream, "</c>");
            OutputStream.Flush();
        }

        private void WriteAsBytes(Stream outStream, string text)
        {
            var bytes = Encoding.UTF8.GetBytes(text);
            outStream.Write(bytes, 0, bytes.Length);
        }

        private void WriteAsBytes(Stream outStream, char[] chars)
        {
            var bytes = Encoding.UTF8.GetBytes(chars);
            outStream.Write(bytes, 0, bytes.Length);
        }
        /**
         * @return  whether the string has leading / trailing spaces that
         *  need to be preserved with the xml:space=\"preserve\" attribute
         */
        private bool HasLeadingTrailingSpaces(string str)
        {
            if (str != null && str.Length > 0)
            {
                char firstChar = str[0];
                char lastChar = str[str.Length - 1];
                return Character.isWhitespace(firstChar) || Character.isWhitespace(lastChar);
            }
            return false;
        }

        //Taken from jdk1.3/src/javax/swing/text/html/HTMLWriter.java
        protected void OutputQuotedString(String s)
        {
            if (s == null || s.Length == 0)
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
                            WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                        }
                        last = counter + 1;

                        WriteAsBytes(OutputStream, "&lt;".ToCharArray());
                        break;
                    case '>':
                        if (counter > last)
                        {
                            WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                        }
                        last = counter + 1;
                        WriteAsBytes(OutputStream, "&gt;".ToCharArray());
                        break;
                    case '&':
                        if (counter > last)
                        {
                            WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                        }
                        last = counter + 1;
                        WriteAsBytes(OutputStream, "&amp;".ToCharArray());
                        break;
                    case '"':
                        if (counter > last)
                        {
                            WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                        }
                        last = counter + 1;
                        WriteAsBytes(OutputStream, "&quot;".ToCharArray());
                        break;
                    // Special characters
                    case '\n':
                    case '\r':
                        if (counter > last)
                        {
                            WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                        }
                        WriteAsBytes(OutputStream, "&#xa;".ToCharArray());
                        last = counter + 1;
                        break;
                    case '\t':
                        if (counter > last)
                        {
                            WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                        }
                        WriteAsBytes(OutputStream, "&#x9;".ToCharArray());
                        last = counter + 1;
                        break;
                    case (char)0xa0:
                        if (counter > last)
                        {
                            WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                        }
                        WriteAsBytes(OutputStream, "&#xa0;".ToCharArray());
                        last = counter + 1;
                        break;
                    default:
                        // YK: XmlBeans silently replaces all ISO control characters ( < 32) with question marks.
                        // the same rule applies to unicode surrogates and "not a character" symbols.
                        if (c < ' ' || Char.IsLowSurrogate(c) || Char.IsHighSurrogate(c) ||
                                ('\uFFFE' <= c && c <= '\uFFFF'))
                        {
                            if (counter > last)
                            {
                                WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                            }
                            WriteAsBytes(OutputStream, "?");
                            last = counter + 1;
                        }
                        else if (c > 127)
                        {
                            if (counter > last)
                            {
                                WriteAsBytes(OutputStream, GetSubArray(chars, last, counter - last));
                            }
                            last = counter + 1;
                            // If the character is outside of UTF8, write the
                            // numeric value.
                            WriteAsBytes(OutputStream, "&#".ToCharArray());
                            WriteAsBytes(OutputStream, ((int)c).ToString());
                            WriteAsBytes(OutputStream, ";");
                        }
                        break;
                }
            }
            if (last < length)
            {
                WriteAsBytes(OutputStream, GetSubArray(chars, last, length - last));
            }
        }

        private char[] GetSubArray(char[] oldArray, int skip, int take)
        {
            var sub = new char[take];
            Array.Copy(oldArray, skip, sub, 0, take);
            return sub;
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
                ret = File.Exists(TemporaryFileInfo.FullName);
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
