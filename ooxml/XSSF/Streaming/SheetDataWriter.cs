using System;
using System.Collections.Generic;
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
            try
            {
                _out.Flush();
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
            ;
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

        void BeginRow(int rownum, SXSSFRow row)
        {
            //TODO: make sure this isn't off.
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
            var text = Encoding.UTF8.GetBytes("<c r=\"" + cellRef + "\"");
            _out.Write(text, 0, text.Length);
            ICellStyle cellStyle = cell.CellStyle;
            if (cellStyle.Index != 0)
            {
                // need to convert the short to unsigned short as the indexes can be up to 64k
                // ideally we would use int for this index, but that would need changes to some more 
                // APIs
                text = Encoding.UTF8.GetBytes(" s=\"" + (cellStyle.Index & 0xffff) + "\"");
                _out.Write(text, 0, text.Length);
            }
            CellType cellType = cell.CellType;
            switch (cellType)
            {
                case CellType.Blank:
                    {
                        text = Encoding.UTF8.GetBytes(">");
                        _out.Write(text, 0, text.Length);
                        break;
                    }
                case CellType.Formula:
                    {
                        //TODO: I may have fucked this up. :)
                        text = Encoding.UTF8.GetBytes(">");
                        _out.Write(text, 0, text.Length);
                        text = Encoding.UTF8.GetBytes("<f>");
                        _out.Write(text, 0, text.Length);
                        outputQuotedString(cell.CellFormula);//THis don't work
                        text = Encoding.UTF8.GetBytes("</f>");
                        _out.Write(text, 0, text.Length);
                        switch (cell.GetCachedFormulaResultTypeEnum())
                        {
                            case CellType.Numeric:
                                double nval = cell.NumericCellValue;
                                if (!Double.IsNaN(nval))
                                {
                                    text = Encoding.UTF8.GetBytes("<v>" + nval + "</v>");
                                    _out.Write(text, 0, text.Length);
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
                            text = Encoding.UTF8.GetBytes(" t=\"" + ST_CellType.s.ToString() + "\">");
                            _out.Write(text, 0, text.Length);
                            text = Encoding.UTF8.GetBytes("<v>");
                            _out.Write(text, 0, text.Length);
                            text = Encoding.UTF8.GetBytes(sRef.ToString());
                            _out.Write(text, 0, text.Length);
                            text = Encoding.UTF8.GetBytes("</v>");
                            _out.Write(text, 0, text.Length);

                        }
                        else
                        {
                            text = Encoding.UTF8.GetBytes(" t=\"inlineStr\">");
                            _out.Write(text, 0, text.Length);
                            text = Encoding.UTF8.GetBytes("<is><t");
                            _out.Write(text, 0, text.Length);

                            if (hasLeadingTrailingSpaces(cell.StringCellValue))
                            {
                                text = Encoding.UTF8.GetBytes(" xml:space=\"preserve\"");
                                _out.Write(text, 0, text.Length);
                                //_out.write(" xml:space=\"preserve\"");
                            }
                            text = Encoding.UTF8.GetBytes(">");
                            _out.Write(text, 0, text.Length);
                            // _out.write(">");
                            outputQuotedString(cell.StringCellValue);//TODO: doesn't work
                            text = Encoding.UTF8.GetBytes("</t></is>");
                            _out.Write(text, 0, text.Length);
                            //_out.write("</t></is>");
                        }
                        break;
                    }
                case CellType.Numeric:
                    {
                        text = Encoding.UTF8.GetBytes(" t=\"n\">");
                        _out.Write(text, 0, text.Length);
                        //_out.write(" t=\"n\">");
                        text = Encoding.UTF8.GetBytes("<v>" + cell.NumericCellValue + "</v>");
                        _out.Write(text, 0, text.Length);
                        //_out.write("<v>" + cell.NumericCellValue + "</v>");
                        break;
                    }
                case CellType.Boolean:
                    {
                        text = Encoding.UTF8.GetBytes(" t=\"b\">");
                        _out.Write(text, 0, text.Length);
                        //_out.write(" t=\"b\">");
                        text = Encoding.UTF8.GetBytes("<v>" + (cell.BooleanCellValue ? "1" : "0") + "</v>");
                        _out.Write(text, 0, text.Length);
                       // _out.write("<v>" + (cell.BooleanCellValue ? "1" : "0") + "</v>");
                        break;
                    }
                case CellType.Error:
                    {
                        FormulaError error = FormulaError.ForInt(cell.ErrorCellValue);

                       // _out.write(" t=\"e\">");
                        text = Encoding.UTF8.GetBytes(" t=\"e\">");
                        _out.Write(text, 0, text.Length);
                        //_out.write("<v>" + error.String + "</v>");
                        text = Encoding.UTF8.GetBytes("<v>" + error.String + "</v>");
                        _out.Write(text, 0, text.Length);
                        break;
                    }
                default:
                    {
                        throw new InvalidOperationException("Invalid cell type: " + cellType);
                    }
            }
            //_out.write("</c>");
            text = Encoding.UTF8.GetBytes("</c>");
            _out.Write(text, 0, text.Length);
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
            throw new NotImplementedException();
        //    if (s == null || s.Length == 0)
        //    {
        //        return;
        //    }

        //    char[]
        //chars = s.ToCharArray();
        //    int last = 0;
        //    int length = s.Length;
        //    for (int counter = 0; counter < length; counter++)
        //    {
        //        char c = chars[counter];
        //        switch (c)
        //        {
        //            case '<':
        //                if (counter > last)
        //                {
        //                    _out.write(chars, last, counter - last);
        //                }
        //                last = counter + 1;
        //                _out.write("&lt;");
        //                break;
        //            case '>':
        //                if (counter > last)
        //                {
        //                    _out.write(chars, last, counter - last);
        //                }
        //                last = counter + 1;
        //                _out.write("&gt;");
        //                break;
        //            case '&':
        //                if (counter > last)
        //                {
        //                    _out.write(chars, last, counter - last);
        //                }
        //                last = counter + 1;
        //                _out.write("&amp;");
        //                break;
        //            case '"':
        //                if (counter > last)
        //                {
        //                    _out.write(chars, last, counter - last);
        //                }
        //                last = counter + 1;
        //                _out.write("&quot;");
        //                break;
        //            // Special characters
        //            case '\n':
        //            case '\r':
        //                if (counter > last)
        //                {
        //                    _out.write(chars, last, counter - last);
        //                }
        //                _out.write("&#xa;");
        //                last = counter + 1;
        //                break;
        //            case '\t':
        //                if (counter > last)
        //                {
        //                    _out.write(chars, last, counter - last);
        //                }
        //                _out.write("&#x9;");
        //                last = counter + 1;
        //                break;
        //            case 0xa0:
        //                if (counter > last)
        //                {
        //                    _out.write(chars, last, counter - last);
        //                }
        //                _out.write("&#xa0;");
        //                last = counter + 1;
        //                break;
        //            default:
        //                // YK: XmlBeans silently replaces all ISO control characters ( < 32) with question marks.
        //                // the same rule applies to unicode surrogates and "not a character" symbols.
        //                if (c < ' ' || Character.isLowSurrogate(c) || Character.isHighSurrogate(c) ||
        //                        ('\uFFFE' <= c && c <= '\uFFFF'))
        //                {
        //                    if (counter > last)
        //                    {
        //                        _out.write(chars, last, counter - last);
        //                    }
        //                    _out.write('?');
        //                    last = counter + 1;
        //                }
        //                else if (c > 127)
        //                {
        //                    if (counter > last)
        //                    {
        //                        _out.write(chars, last, counter - last);
        //                    }
        //                    last = counter + 1;
        //                    // If the character is outside of UTF8, write the
        //                    // numeric value.
        //                    _out.write("&#");
        //                    _out.write(((int)c).ToString());
        //                    _out.write(";");
        //                }
        //                break;
        //        }
        //    }
        //    if (last < length)
        //    {
        //        _out.write(chars, last, length - last);
        //    }
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
