/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Text;

namespace NPOI.XSSF.EventUserModel
{

    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using NPOI.XSSF.Binary;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// Reader for xlsb files.
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBReader : XSSFReader
    {

        //private static  POILogger log = POILogFactory.GetLogger(XSSFBReader.class);
        private static  ISet<String> WORKSHEET_RELS =
            (new HashSet<String>(
                    Arrays.AsList(new String[]{
                            XSSFRelation.WORKSHEET.Relation,
                            XSSFRelation.CHARTSHEET.Relation,
                            XSSFRelation.MACRO_SHEET_BIN.Relation,
                            XSSFRelation.INTL_MACRO_SHEET_BIN.Relation,
                            XSSFRelation.DIALOG_SHEET_BIN.Relation
                    })
            ));

        /// <summary>
        /// Creates a new XSSFReader, for the given package
        /// </summary>
        /// <param name="pkg">opc package</param>
        public XSSFBReader(OPCPackage pkg)
            : base(pkg)
        {


        }

        /// <summary>
        /// In Excel 2013, the absolute path where the file was last saved may be stored in
        /// the <see cref="XSSFBRecordType.BrtAbsPath15" /> record.  The equivalent in ooxml is
        /// &lt;x15ac:absPath&gt;.
        /// </summary>
        /// <return>absolute path or <c>null</c> if it could not be found.</return>
        /// <exception cref="IOException">when there's a problem with the workbook part's stream</exception>
        public String GetAbsPathMetadata()
        {

            Stream is1 = null;
            try
            {
                is1 = workbookPart.GetInputStream();
                PathExtractor p = new PathExtractor(workbookPart.GetInputStream());
                p.Parse();
                return p.GetPath();
            }
            finally
            {
                IOUtils.CloseQuietly(is1);
            }
        }

        /// <summary>
        /// Returns an Iterator which will let you Get at all the
        ///  different Sheets in turn.
        /// Each sheet's InputStream is only opened when fetched
        ///  from the Iterator. It's up to you to close the
        ///  InputStreams when done with each one.
        /// </summary>
        public override IEnumerator<Stream> GetSheetsData()
        {
            return new SheetIterator(workbookPart);
        }

        public XSSFBStylesTable GetXSSFBStylesTable()
        {

            List<PackagePart> parts = pkg.GetPartsByContentType(XSSFBRelation.STYLES_BINARY.ContentType);
            if(parts.Count == 0)
                return null;

            // Create the Styles Table, and associate the Themes if present
            return new XSSFBStylesTable(parts[0].GetInputStream());

        }

        public class SheetIterator : XSSFReader.SheetIterator
        {

            /// <summary>
            /// Construct a new SheetIterator
            /// </summary>
            /// <param name="wb">package part holding workbook.xml</param>
            public SheetIterator(PackagePart wb)
            : base(wb)
            {


            }
            public static ISet<String> GetSheetRelationships()
            {
                return WORKSHEET_RELS;
            }

            public override List<XSSFSheetRef> CreateSheetIteratorFromWB(PackagePart wb)
            {

                SheetRefLoader sheetRefLoader = new SheetRefLoader(wb.GetInputStream());
                sheetRefLoader.Parse();
                return sheetRefLoader.GetSheets();
            }

            /// <summary>
            /// Not supported by XSSFBReader's SheetIterator.
            /// Please use <see cref="getXSSFBSheetComments()" /> instead.
            /// </summary>
            /// <return>nothing, always throws ArgumentException!</return>
            public CommentsTable GetSheetComments()
            {
                throw new ArgumentException("Please use GetXSSFBSheetComments");
            }

            public XSSFBCommentsTable GetXSSFBSheetComments()
            {
                PackagePart sheetPkg = SheetPart;

                // Do we have a comments relationship? (Only ever one if so)
                try
                {
                    PackageRelationshipCollection commentsList =
                        sheetPkg.GetRelationshipsByType(XSSFRelation.SHEET_COMMENTS.Relation);
                    if(commentsList.Size > 0)
                    {
                        PackageRelationship comments = commentsList.GetRelationship(0);
                        if(comments == null || comments.TargetUri == null)
                        {
                            return null;
                        }
                        PackagePartName commentsName = PackagingUriHelper.CreatePartName(comments.TargetUri);
                        PackagePart commentsPart = sheetPkg.Package.GetPart(commentsName);
                        return new XSSFBCommentsTable(commentsPart.GetInputStream());
                    }
                }
                catch(InvalidFormatException e)
                {
                    return null;
                }
                catch(IOException e)
                {
                    return null;
                }
                return null;
            }

        }


        private class PathExtractor : XSSFBParser
        {
            private static BitArray RECORDS = new BitArray(2176);
            static PathExtractor()
            {
                RECORDS.Set((int) XSSFBRecordType.BrtAbsPath15, true);
            }
            private String path = null;
            public PathExtractor(Stream is1)
            : base(is1, RECORDS)
            {

            }
            public override void HandleRecord(int recordType, byte[] data)
            {

                if(recordType != (int) XSSFBRecordType.BrtAbsPath15)
                {
                    return;
                }
                StringBuilder sb = new StringBuilder();
                XSSFBUtils.ReadXLWideString(data, 0, sb);
                path = sb.ToString();
            }

            /// <summary>
            /// </summary>
            /// <return>the path if found, otherwise <c>null</c></return>
            public String GetPath()
            {
                return path;
            }
        }

        private class SheetRefLoader : XSSFBParser
        {
            List<XSSFSheetRef> sheets = new List<XSSFSheetRef>();

            public SheetRefLoader(Stream is1)
            : base(is1)
            {

            }
            public override void HandleRecord(int recordType, byte[] data)
            {

                if(recordType == (int) XSSFBRecordType.BrtBundleSh)
                {
                    addWorksheet(data);
                }
            }

            private void addWorksheet(byte[] data)
            {
                //try to parse the BrtBundleSh
                //if there's an exception, catch it and
                //try to figure out if this is one of the old beta-created xlsb files
                //or if this is a general exception
                try
                {
                    tryToAddWorksheet(data);
                }
                catch(XSSFBParseException e)
                {
                    if(tryOldFormat(data))
                    {
                        //log.log(POILogger.WARN, "This file was written with a beta version of Excel. "+
                        //        "POI will try to parse the file as a regular xlsb.");
                    }
                    else
                    {
                        throw;
                    }
                }
            }

            private void tryToAddWorksheet(byte[] data)
            {

                int offset = 0;
                //this is the sheet state #2.5.142
                long hsShtat = LittleEndian.GetUInt(data, offset);
                offset += LittleEndian.INT_SIZE;

                long iTabID = LittleEndian.GetUInt(data, offset);
                offset += LittleEndian.INT_SIZE;
                //according to #2.4.304
                if(iTabID < 1 || iTabID > 0x0000FFFFL)
                {
                    throw new XSSFBParseException("table id out of range: "+iTabID);
                }
                StringBuilder sb = new StringBuilder();
                offset += XSSFBUtils.ReadXLWideString(data, offset, sb);
                String relId = sb.ToString();
                sb.Length = (0);
                offset += XSSFBUtils.ReadXLWideString(data, offset, sb);
                String name = sb.ToString();
                if(relId != null && relId.Trim().Length > 0)
                {
                    sheets.Add(new XSSFSheetRef(relId, name));
                }
            }

            private bool tryOldFormat(byte[] data)
            {

                //undocumented what is contained in these 8 bytes.
                //for the non-beta xlsb files, this would be 4, not 8.
                int offset = 8;
                long iTabID = LittleEndian.GetUInt(data, offset);
                offset += LittleEndian.INT_SIZE;
                if(iTabID < 1 || iTabID > 0x0000FFFFL)
                {
                    throw new XSSFBParseException("table id out of range: "+iTabID);
                }
                StringBuilder sb = new StringBuilder();
                offset += XSSFBUtils.ReadXLWideString(data, offset, sb);
                String relId = sb.ToString();
                sb.Length = (0);
                offset += XSSFBUtils.ReadXLWideString(data, offset, sb);
                String name = sb.ToString();
                if(relId != null && relId.Trim().Length > 0)
                {
                    sheets.Add(new XSSFSheetRef(relId, name));
                }
                if(offset == data.Length)
                {
                    return true;
                }
                return false;
            }

            public List<XSSFSheetRef> GetSheets()
            {
                return sheets;
            }
        }
    }
}

