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

namespace NPOI.HSSF.Extractor
{
    using System;
    using System.IO;
    using System.Text; 
using Cysharp.Text;
    using NPOI.HSSF;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    /**
     * A text extractor for old Excel files, which are too old for
     *  HSSFWorkbook to handle. This includes Excel 95, and very old 
     *  (pre-OLE2) Excel files, such as Excel 4 files.
     * <p>
     * Returns much (but not all) of the textual content of the file, 
     *  suitable for indexing by something like Apache Lucene, or used
     *  by Apache Tika, but not really intended for display to the user.
     * </p>
     */
    public class OldExcelExtractor
    {
        private const int FILE_PASS_RECORD_SID = 0x2f;
        private RecordInputStream ris;

        // sometimes we hold the stream here and thus need to ensure it is closed at some point
        private ICloseable toClose;
        private Stream toCloseStream;

        private int biffVersion;
        private int fileType;

        public OldExcelExtractor(Stream input)
        {
            Open(input);
        }
        public OldExcelExtractor(FileInfo f)
        {
            NPOIFSFileSystem poifs = null;
            try
            {
                poifs = new NPOIFSFileSystem(f);
                toClose = poifs;
                Open(poifs);
                return;
            }
            catch (OldExcelFormatException)
            {
                // will be handled by workaround below
                if (poifs != null)
                {
                    poifs.Close();
                }
            }
            catch (NotOLE2FileException)
            {
                // will be handled by workaround below
                if (poifs != null)
                {
                    poifs.Close();
                }
            }

            FileStream biffStream = f.OpenRead();
            try
            {
                Open(biffStream);
            }
            catch (IOException)
            {
                // ensure that the stream is properly closed here if an Exception
                // is thrown while opening
                biffStream.Close();
                throw;
            }
            catch (RuntimeException)
            {
                // ensure that the stream is properly closed here if an Exception
                // is thrown while opening
                biffStream.Close();
                throw;
            }
        }
        public OldExcelExtractor(NPOIFSFileSystem fs)
        {
            Open(fs);
        }
        public OldExcelExtractor(DirectoryNode directory)
        {
            Open(directory);
        }

        private void Open(Stream biffStream)
        {
            BufferedStream bis = (biffStream is BufferedStream stream) 
            ? stream
            : new BufferedStream(biffStream, 8);

            if (NPOIFSFileSystem.HasPOIFSHeader(bis))
            {
                NPOIFSFileSystem poifs = new NPOIFSFileSystem(bis);
                try
                {
                    Open(poifs);
                }
                finally
                {
                    poifs.Close();
                }
            }
            else
            {
                ris = new RecordInputStream(bis);
                toCloseStream = bis;
                Prepare();
            }
        }
        private void Open(NPOIFSFileSystem fs)
        {
            Open(fs.Root);
        }
        private void Open(DirectoryNode directory)
        {
            DocumentNode book;
            try
            {
                book = (DocumentNode)directory.GetEntry(InternalWorkbook.OLD_WORKBOOK_DIR_ENTRY_NAME);
            }
            catch (FileNotFoundException)
            {
                // some files have "Workbook" instead
                book = (DocumentNode)directory.GetEntry(InternalWorkbook.WORKBOOK_DIR_ENTRY_NAMES[0]);
            }
            if (book == null)
            {
                throw new IOException("No Excel 5/95 Book stream found");
            }

            ris = new RecordInputStream(directory.CreateDocumentInputStream(book));
            Prepare();
        }

        public static void main(String[] args)
        {
            if (args.Length < 1)
            {
                System.Console.WriteLine("Use:");
                System.Console.WriteLine("   OldExcelExtractor <filename>");
                return;
            }
            OldExcelExtractor extractor = new OldExcelExtractor(new FileInfo(args[0]));
            System.Console.WriteLine(extractor.Text);
            extractor.Close();
        }

        private void Prepare()
        {
            if (!ris.HasNextRecord)
                throw new ArgumentException("File Contains no records!");
            ris.NextRecord();

            // Work out what version we're dealing with
            int bofSid = ris.Sid;
            switch (bofSid)
            {
                case BOFRecord.biff2_sid:
                    biffVersion = 2;
                    break;
                case BOFRecord.biff3_sid:
                    biffVersion = 3;
                    break;
                case BOFRecord.biff4_sid:
                    biffVersion = 4;
                    break;
                case BOFRecord.biff5_sid:
                    biffVersion = 5;
                    break;
                default:
                    throw new ArgumentException("File does not begin with a BOF, found sid of " + bofSid);
            }

            // Get the type
            BOFRecord bof = new BOFRecord(ris);
            fileType = (int)bof.Type;
        }

        /**
         * The Biff version, largely corresponding to the Excel version
         */
        public int BiffVersion
        {
            get
            {
                return biffVersion;
            }
        }
        /**
         * The kind of the file, one of {@link BOFRecord#TYPE_WORKSHEET},
         *  {@link BOFRecord#TYPE_CHART}, {@link BOFRecord#TYPE_EXCEL_4_MACRO}
         *  or {@link BOFRecord#TYPE_WORKSPACE_FILE}
         */
        public int FileType
        {
            get
            {
                return fileType;
            }
        }

        /**
         * Retrieves the text contents of the file, as best we can
         *  for these old file formats
         */
        public String Text
        {
            get
            {
                StringBuilder text = new StringBuilder();

                // To track formats and encodings
                CodepageRecord codepage = null;
                // TODO track the XFs and Format Strings

                // Process each record in turn, looking for interesting ones
                while (ris.HasNextRecord)
                {
                    int sid = ris.GetNextSid();
                    ris.NextRecord();

                    switch (sid)
                    {
                        case FILE_PASS_RECORD_SID:
                            throw new EncryptedDocumentException("Encryption not supported for Old Excel files");

                        // Biff 5+ only, no sheet names in older formats
                        case OldSheetRecord.sid:
                            OldSheetRecord shr = new OldSheetRecord(ris);
                            shr.SetCodePage(/*setter*/codepage);
                            text.Append("Sheet: ");
                            text.Append(shr.Sheetname);
                            text.Append('\n');
                            break;

                        case OldLabelRecord.biff2_sid:
                        case OldLabelRecord.biff345_sid:
                            OldLabelRecord lr = new OldLabelRecord(ris);
                            lr.SetCodePage(/*setter*/codepage);
                            text.Append(lr.Value);
                            text.Append('\n');
                            break;
                        case OldStringRecord.biff2_sid:
                        case OldStringRecord.biff345_sid:
                            OldStringRecord sr = new OldStringRecord(ris);
                            sr.SetCodePage(/*setter*/codepage);
                            text.Append(sr.GetString());
                            text.Append('\n');
                            break;

                        case NumberRecord.sid:
                            NumberRecord nr = new NumberRecord(ris);
                            handleNumericCell(text, nr.Value);
                            break;
                        case OldFormulaRecord.biff2_sid:
                        case OldFormulaRecord.biff3_sid:
                        case OldFormulaRecord.biff4_sid:
                            // Biff 2 and 5+ share the same SID, due to a bug...
                            if (biffVersion == 5)
                            {
                                FormulaRecord fr = new FormulaRecord(ris);
                                if (fr.CachedResultType == CellType.Numeric)
                                {
                                    handleNumericCell(text, fr.Value);
                                }
                            }
                            else
                            {
                                OldFormulaRecord fr = new OldFormulaRecord(ris);
                                if (fr.GetCachedResultType() == CellType.Numeric)
                                {
                                    handleNumericCell(text, fr.Value);
                                }
                            }
                            break;
                        case RKRecord.sid:
                            RKRecord rr = new RKRecord(ris);
                            handleNumericCell(text, rr.RKNumber);
                            break;

                        case CodepageRecord.sid:
                            codepage = new CodepageRecord(ris);
                            break;

                        default:
                            ris.ReadFully(new byte[ris.Remaining]);
                            break;
                    }
                }

                Close();
                ris = null;
                return text.ToString();
            }
        }

        protected void handleNumericCell(StringBuilder text, double value)
        {
            // TODO Need to fetch / use format strings
            text.Append(value);
            text.Append('\n');
        }

        public void Close()
        {
            // some cases require this close here
            if (toClose != null)
            {
                IOUtils.CloseQuietly(toClose);
                toClose = null;
            }
            if (toCloseStream != null)
            {
                IOUtils.CloseQuietly(toCloseStream);
                toClose = null;
            }
        }
    }

}