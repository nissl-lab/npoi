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

namespace NPOI.XSSF.Binary
{
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;

    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBSharedStringsTable
    {

        /// <summary>
        /// An integer representing the total count of strings in the workbook. This count does not
        /// include any numbers, it counts only the total of text strings in the workbook.
        /// </summary>
        private int count;

        /// <summary>
        /// An integer representing the total count of unique strings in the Shared String Table.
        /// A string is unique even if it is a copy of another string, but has different formatting applied
        /// at the character level.
        /// </summary>
        private int uniqueCount;

        /// <summary>
        /// The shared strings table.
        /// </summary>
        private List<String> strings = new List<String>();

        /// <summary>
        /// </summary>
        /// <param name="pkg">The <see cref="OPCPackage"/> to use as basis for the shared-strings table.</param>
        /// <exception cref="IOException">If reading the data from the package Assert.Fails.</exception>
        /// <exception cref="SAXException">if parsing the XML data Assert.Fails.</exception>
        public XSSFBSharedStringsTable(OPCPackage pkg)
        {
            List<PackagePart> parts =
                pkg.GetPartsByContentType(XSSFBRelation.SHARED_STRINGS_BINARY.ContentType);

            // Some workbooks have no shared strings table.
            if(parts.Count > 0)
            {
                PackagePart sstPart = parts[0];

                readFrom(sstPart.GetInputStream());
            }
        }

        /// <summary>
        /// Like POIXMLDocumentPart constructor
        /// </summary>
        public XSSFBSharedStringsTable(PackagePart part)
        {

            readFrom(part.GetInputStream());
        }

        private void readFrom(Stream inputStream)
        {

            SSTBinaryReader reader = new SSTBinaryReader(inputStream, this);
            reader.Parse();
        }

        /// <summary>
        /// </summary>
        /// <return>a defensive copy of strings</return>
        public List<String> GetItems()
        {
            List<String> ret = new List<String>(strings.Count);
            ret.AddRange(strings);
            return ret;
        }

        public String GetEntryAt(int i)
        {
            return strings[i];
        }

        /// <summary>
        /// Return an integer representing the total count of strings in the workbook. This count does not
        /// include any numbers, it counts only the total of text strings in the workbook.
        /// </summary>
        /// <return>the total count of strings in the workbook</return>
        public int GetCount()
        {
            return this.count;
        }

        /// <summary>
        /// Returns an integer representing the total count of unique strings in the Shared String Table.
        /// A string is unique even if it is a copy of another string, but has different formatting applied
        /// at the character level.
        /// </summary>
        /// <return>the total count of unique strings in the workbook</return>
        public int GetUniqueCount()
        {
            return this.uniqueCount;
        }

        private class SSTBinaryReader : XSSFBParser
        {
            private XSSFBSharedStringsTable sst;
            public SSTBinaryReader(Stream is1, XSSFBSharedStringsTable sst)
            : base(is1)
            {
                this.sst = sst;
            }
            public override void HandleRecord(int recordType, byte[] data)
            {

                XSSFBRecordType type = XSSFBRecordTypeClass.Lookup(recordType);

                switch(type)
                {
                    case XSSFBRecordType.BrtSstItem:
                        XSSFBRichStr rstr = XSSFBRichStr.Build(data, 0);
                        sst.strings.Add(rstr.String);
                        break;
                    case XSSFBRecordType.BrtBeginSst:
                        sst.count = XSSFBUtils.CastToInt(LittleEndian.GetUInt(data, 0));
                        sst.uniqueCount = XSSFBUtils.CastToInt(LittleEndian.GetUInt(data, 4));
                        break;
                }

            }
        }

    }
}

