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

        public FileInfo _fd;
        public FileStream _out;
        public int _rownum;
        public int _numberOfFlushedRows;
        public int _lowestIndexOfFlushedRows; // meaningful only of _numberOfFlushedRows>0
        public int _numberOfCellsOfLastFlushedRow; // meaningful only of _numberOfFlushedRows>0
        public int _numberLastFlushedRow = -1; // meaningful only of _numberOfFlushedRows>0

        /**
 * Table of strings shared across this workbook.
 * If two cells contain the same string, then the cell value is the same index into SharedStringsTable
 */
        private SharedStringsTable _sharedStringSource;

        public SheetDataWriter()
        {
            _fd = createTempFile();
            _out = createWriter(_fd);
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
        //TODO: may want to use a different writer object
        public FileInfo createTempFile()
        {
            return TempFile.CreateTempFile("poi-sxssf-sheet", ".xml");
        }

        /**
         * Create a writer for the sheet data.
         * 
         * @param  fd the file to write to
         */
        public FileStream createWriter(FileInfo fd)
        {

            FileStream fos = null;
            FileStream decorated;
            try
            {
                fos = new FileStream(fd.FullName, FileMode.Append, FileAccess.Write);
                //decorated = decorateOutputStream(fos);
            }
            catch (Exception e)
            {
                if (fos != null)
                {
                    fos.Close();
                }

                throw e;
            }
            //TODO: this is the decorate?
            //StreamWriter sw = new StreamWriter(fos, Encoding.UTF8);
            return fos;
            //return new BufferedStream(
            //        new BinaryWriter(fos, Encoding.UTF8).BaseStream);
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
        //TODO: may not need to decorate in C#
        protected Stream decorateOutputStream(Stream fos)
        {
            return fos;
        }

        /**
         * flush and close the temp data writer. 
         * This method <em>must</em> be invoked before calling {@link #getWorksheetXMLInputStream()}
         */
        public void Close()
        {
            //TODO: test
            try
            {
                _out.Flush(true);
            }
            catch (Exception e)
            {

            }
            try
            {
                _out.Close();
            }
            catch (Exception e)
            {

            }


        }

        /**
         * @return a stream to read temp file with the sheet data
         */
        public Stream GetWorksheetXMLInputStream()
        {
            Stream fis = new FileStream(_fd.FullName, FileMode.Open, FileAccess.ReadWrite);
            try
            {
                return decorateInputStream(fis);
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
        protected Stream decorateInputStream(Stream fis)
        {
            return fis;
        }



        protected void Finalize()
        {
            _fd.Delete();
            if (File.Exists(_fd.FullName))
            {
                logger.Log(POILogger.ERROR, "Can't delete temporary encryption file: " + _fd);
            }
            //TODO: accomplish whatever this does.
            //super.finalize();
        }

        /**
         * Write a row to the file
         *
         * @param rownum 0-based row number
         * @param row    a row
         */
        public void WriteRow(int rownum, SXSSFRow row)
        {
            if (_numberOfFlushedRows == 0)
                _lowestIndexOfFlushedRows = rownum;
            _numberLastFlushedRow = Math.Max(rownum, _numberLastFlushedRow);
            _numberOfCellsOfLastFlushedRow = row.LastCellNum;
            _numberOfFlushedRows++;
            BeginRow(rownum, row);
            var cells = row.GetEnumerator();
            int columnIndex = 0;
            while (cells.MoveNext())
            {
                writeCell(columnIndex++, cells.Current);
            }
            endRow();
        }

        private void BeginRow(int rownum, SXSSFRow row)
        {
            var text = Encoding.UTF8.GetBytes("<row r=\"" + (rownum + 1) + "\"");
            _out.Write(text, 0, text.Length);
            if (row.hasCustomHeight())
            {
                text = Encoding.UTF8.GetBytes(" customHeight=\"true\"  ht=\"" + row.HeightInPoints + "\"");
                _out.Write(text, 0, text.Length);
            }
            if (row.ZeroHeight)
            {
                text = Encoding.UTF8.GetBytes(" hidden=\"true\"");
                _out.Write(text, 0, text.Length);
            }
            if (row.IsFormatted)
            {
                text = Encoding.UTF8.GetBytes(" s=\"" + row.RowStyle.Index + "\"");
                _out.Write(text, 0, text.Length);
                text = Encoding.UTF8.GetBytes(" customFormat=\"1\"");
                _out.Write(text, 0, text.Length);
            }
            //TODO: _outlinelevel or OUTLINE LEVEL
            if (row.OutlineLevel != 0)
            {
                text = Encoding.UTF8.GetBytes(" outlineLevel=\"" + row._outlineLevel + "\"");
                _out.Write(text, 0, text.Length);
            }
            if (row._hidden != null)
            {
                text = Encoding.UTF8.GetBytes(" hidden=\"" + (row._hidden.Value ? "1" : "0") + "\"");
                _out.Write(text, 0, text.Length);
            }
            if (row._collapsed != null)
            {
                text = Encoding.UTF8.GetBytes(" collapsed=\"" + (row._collapsed.Value ? "1" : "0") + "\"");
                _out.Write(text, 0, text.Length);
            }

            text = Encoding.UTF8.GetBytes(">\n");
            _out.Write(text, 0, text.Length);
            this._rownum = rownum;
        }

        void endRow()
        {
            var text = Encoding.UTF8.GetBytes("</row>\n");
            _out.Write(text, 0, text.Length);
        }

        //TODO: The strings that need to be written are probably wrong. :\
        public void writeCell(int columnIndex, ICell cell)
        {
            if (cell == null)
            {
                return;
            }
            string cellRef = new CellReference(_rownum, columnIndex).FormatAsString();
            WriteAsBytes(_out, "<c r=\"" + cellRef + "\"");
            ICellStyle cellStyle = cell.CellStyle;
            if (cellStyle.Index != 0)
            {
                // need to convert the short to unsigned short as the indexes can be up to 64k
                // ideally we would use int for this index, but that would need changes to some more 
                // APIs
                WriteAsBytes(_out, " s=\"" + (cellStyle.Index & 0xffff) + "\"");
            }
            CellType cellType = cell.CellType;
            switch (cellType)
            {
                case CellType.Blank:
                    {
                        WriteAsBytes(_out, ">");
                        break;
                    }
                case CellType.Formula:
                    {
                        WriteAsBytes(_out, ">");
                        WriteAsBytes(_out, "<f>");

                        outputQuotedString(cell.CellFormula);//TODO: this doesn't work

                        WriteAsBytes(_out, "</f>");

                        switch (cell.GetCachedFormulaResultTypeEnum())
                        {
                            case CellType.Numeric:
                                double nval = cell.NumericCellValue;
                                if (!Double.IsNaN(nval))
                                {
                                    WriteAsBytes(_out, "<v>" + nval + "</v>");
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

                            //TODO: is this supposed to be and s=\"
                            WriteAsBytes(_out, " t=\"" + ST_CellType.s.ToString() + "\">");
                            WriteAsBytes(_out, "<v>");
                            WriteAsBytes(_out, sRef.ToString());
                            WriteAsBytes(_out, "</v>");
                        }
                        else
                        {
                            WriteAsBytes(_out, " t=\"inlineStr\">");
                            WriteAsBytes(_out, "<is><t");

                            if (hasLeadingTrailingSpaces(cell.StringCellValue))
                            {
                                WriteAsBytes(_out, " xml:space=\"preserve\"");
                            }

                            WriteAsBytes(_out, ">");

                            outputQuotedString(cell.StringCellValue);//TODO: doesn't work

                            WriteAsBytes(_out, "</t></is>");
                        }
                        break;
                    }
                case CellType.Numeric:
                    {
                        WriteAsBytes(_out, " t=\"n\">");
                        WriteAsBytes(_out, "<v>" + cell.NumericCellValue + "</v>");
                        break;
                    }
                case CellType.Boolean:
                    {
                        WriteAsBytes(_out, " t=\"b\">");
                        WriteAsBytes(_out, "<v>" + (cell.BooleanCellValue ? "1" : "0") + "</v>");
                        break;
                    }
                case CellType.Error:
                    {
                        FormulaError error = FormulaError.ForInt(cell.ErrorCellValue);

                        WriteAsBytes(_out, " t=\"e\">");
                        WriteAsBytes(_out, "<v>" + error.String + "</v>");
                        break;
                    }
                default:
                    {
                        throw new InvalidOperationException("Invalid cell type: " + cellType);
                    }
            }
            WriteAsBytes(_out, "</c>");
            _out.Flush();
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
        bool hasLeadingTrailingSpaces(string str)
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
        protected void outputQuotedString(String s)
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
                            WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                            //_out.Write(chars, last, counter - last);
                        }
                        last = counter + 1;
                      
                        WriteAsBytes(_out, "&lt;".ToCharArray());
                       // _out.write("&lt;");
                        break;
                    case '>':
                        if (counter > last)
                        {
                            WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                            //_out.write(chars, last, counter - last);
                        }
                        last = counter + 1;
                        WriteAsBytes(_out, "&gt;".ToCharArray());
                       // _out.write("&gt;");
                        break;
                    case '&':
                        if (counter > last)
                        {
                            WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                            //_out.write(chars, last, counter - last);
                        }
                        last = counter + 1;
                        WriteAsBytes(_out, "&amp;".ToCharArray());
                        //_out.write("&amp;");
                        break;
                    case '"':
                        if (counter > last)
                        {
                            WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                           // _out.write(chars, last, counter - last);
                        }
                        last = counter + 1;
                        WriteAsBytes(_out, "&quot;".ToCharArray());
                        //_out.write("&quot;");
                        break;
                    // Special characters
                    case '\n':
                    case '\r':
                        if (counter > last)
                        {
                            WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                            //_out.write(chars, last, counter - last);
                        }
                        WriteAsBytes(_out, "&#xa;".ToCharArray());
                        //_out.write("&#xa;");
                        last = counter + 1;
                        break;
                    case '\t':
                        if (counter > last)
                        {
                            WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                            //_out.write(chars, last, counter - last);
                        }
                        WriteAsBytes(_out, "&#x9;".ToCharArray());
                        //_out.write("&#x9;");
                        last = counter + 1;
                        break;
                    case (char)0xa0:
                        if (counter > last)
                        {
                            WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                           // _out.write(chars, last, counter - last);
                        }
                        WriteAsBytes(_out, "&#xa0;".ToCharArray());
                        //_out.write("&#xa0;");
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
                                WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                                //_out.write(chars, last, counter - last);
                            }
                            WriteAsBytes(_out, "?");
                           // _out.write('?');
                            last = counter + 1;
                        }
                        else if (c > 127)
                        {
                            if (counter > last)
                            {
                                WriteAsBytes(_out, chars.Skip(last).Take(counter - last).ToArray());
                               // _out.write(chars, last, counter - last);
                            }
                            last = counter + 1;
                            // If the character is outside of UTF8, write the
                            // numeric value.
                            WriteAsBytes(_out, "&#".ToCharArray());
                            //_out.write("&#");
                            WriteAsBytes(_out, ((int)c).ToString());
                            //_out.write(((int)c).ToString());
                            WriteAsBytes(_out, ";");
                            //_out.write(";");
                        }
                        break;
                }
            }
            if (last < length)
            {
                WriteAsBytes(_out, chars.Skip(last).Take(length - last).ToArray());
               // _out.write(chars, last, length - last);
            }
        }

        /**
         * Deletes the temporary file that backed this sheet on disk.
         * @return true if the file was deleted, false if it wasn't.
         */
        public bool dispose()
        {
            bool ret;
            try
            {
                _out.Close();
            }
            finally
            {
                _fd.Delete();
                ret = File.Exists(_fd.FullName);
            }
            return ret;
        }
    }
}
