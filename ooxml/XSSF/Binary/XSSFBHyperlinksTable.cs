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
using System.Runtime.Serialization;
using System.Text;

namespace NPOI.XSSF.Binary
{

    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.Util;
    using NPOI.SS.Util;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFBHyperlinksTable
    {

        private static  BitArray RECORDS = new BitArray(2176);


        static XSSFBHyperlinksTable()
        {
            RECORDS.Set((int)XSSFBRecordType.BrtHLink, true);
        }


        private  List<XSSFHyperlinkRecord> hyperlinkRecords = new List<XSSFHyperlinkRecord>();

        //cache the relId to hyperlink url from the sheet's .rels
        private Dictionary<String, String> relIdToHyperlink = new Dictionary<String, String>();

        public XSSFBHyperlinksTable(PackagePart sheetPart)
        {

            //load the urls from the sheet .rels
            loadUrlsFromSheetRels(sheetPart);
            //now load the hyperlinks from the bottom of the sheet
            HyperlinkSheetScraper scraper = new HyperlinkSheetScraper(sheetPart.GetInputStream(), this);
            scraper.Parse();
        }

        /// <summary>
        /// </summary>
        /// <return>a map of the hyperlinks. The key is the top left cell address in their CellRange</return>
        public IDictionary<CellAddress, List<XSSFHyperlinkRecord>> GetHyperLinks()
        {
            SortedDictionary<CellAddress, List<XSSFHyperlinkRecord>> hyperlinkMap =
                new SortedDictionary<CellAddress, List<XSSFHyperlinkRecord>>(new TopLeftCellAddressComparator());
            foreach(XSSFHyperlinkRecord hyperlinkRecord in hyperlinkRecords)
            {
                CellAddress cellAddress = new CellAddress(hyperlinkRecord.CellRangeAddress.FirstRow,
                    hyperlinkRecord.CellRangeAddress.FirstColumn);
                hyperlinkMap.TryGetValue(cellAddress, out List<XSSFHyperlinkRecord> list);
                if(list == null)
                {
                    list = new List<XSSFHyperlinkRecord>();
                }
                list.Add(hyperlinkRecord);
                hyperlinkMap[cellAddress] = list;
            }
            return hyperlinkMap;
        }


        /// <summary>
        /// </summary>
        /// <param name="cellAddress">cell address to find</param>
        /// <return>null if not a hyperlink</return>
        public List<XSSFHyperlinkRecord> findHyperlinkRecord(CellAddress cellAddress)
        {
            List<XSSFHyperlinkRecord> overlapping = null;
            CellRangeAddress targetCellRangeAddress = new CellRangeAddress(cellAddress.Row,
                cellAddress.Row,
                cellAddress.Column,
                cellAddress.Column);
            foreach(XSSFHyperlinkRecord record in hyperlinkRecords)
            {
                if(CellRangeUtil.Intersect(targetCellRangeAddress, record.CellRangeAddress) != CellRangeUtil.NO_INTERSECTION)
                {
                    if(overlapping == null)
                    {
                        overlapping = new List<XSSFHyperlinkRecord>();
                    }
                    overlapping.Add(record);
                }
            }
            return overlapping;
        }

        private void loadUrlsFromSheetRels(PackagePart sheetPart)
        {
            try
            {
                foreach(PackageRelationship rel in sheetPart.GetRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.Relation))
                {
                    relIdToHyperlink[rel.Id] = rel.TargetUri.ToString();
                }
            }
            catch(InvalidFormatException e)
            {
                //swallow
            }
        }

        private class HyperlinkSheetScraper : XSSFBParser
        {

            private XSSFBCellRange hyperlinkCellRange = new XSSFBCellRange();
            private  StringBuilder xlWideStringBuilder = new StringBuilder();
            private XSSFBHyperlinksTable hlt;
            public HyperlinkSheetScraper(Stream is1, XSSFBHyperlinksTable hlt)
            : base(is1, RECORDS)
            {
                this.hlt = hlt;
            }
            public override void HandleRecord(int recordType, byte[] data)
            {

                if(recordType != (int) XSSFBRecordType.BrtHLink)
                {
                    return;
                }
                int offset = 0;
                String relId = "";
                String location = "";
                String toolTip = "";
                String display = "";

                hyperlinkCellRange = XSSFBCellRange.parse(data, offset, hyperlinkCellRange);
                offset += XSSFBCellRange.Length;
                xlWideStringBuilder.Length = (0);
                offset += XSSFBUtils.ReadXLNullableWideString(data, offset, xlWideStringBuilder);
                relId = xlWideStringBuilder.ToString();
                xlWideStringBuilder.Length = (0);
                offset += XSSFBUtils.ReadXLWideString(data, offset, xlWideStringBuilder);
                location = xlWideStringBuilder.ToString();
                xlWideStringBuilder.Length = (0);
                offset += XSSFBUtils.ReadXLWideString(data, offset, xlWideStringBuilder);
                toolTip = xlWideStringBuilder.ToString();
                xlWideStringBuilder.Length = (0);
                offset += XSSFBUtils.ReadXLWideString(data, offset, xlWideStringBuilder);
                display = xlWideStringBuilder.ToString();
                CellRangeAddress cellRangeAddress = new CellRangeAddress(hyperlinkCellRange.firstRow, hyperlinkCellRange.lastRow, hyperlinkCellRange.firstCol, hyperlinkCellRange.lastCol);

                hlt.relIdToHyperlink.TryGetValue(relId, out string url);
                if(location == null || location.Length == 0)
                {
                    location = url;
                }

                hlt.hyperlinkRecords.Add(
                        new XSSFHyperlinkRecord(cellRangeAddress, relId, location, toolTip, display)
                );
            }
        }

        internal class TopLeftCellAddressComparator : IComparer<CellAddress>
        {
            private static  long serialVersionUID = 1L;
            public int Compare(CellAddress o1, CellAddress o2)
            {
                if(o1.Row < o2.Row)
                {
                    return -1;
                }
                else if(o1.Row > o2.Row)
                {
                    return 1;
                }
                if(o1.Column < o2.Column)
                {
                    return -1;
                }
                else if(o1.Column > o2.Column)
                {
                    return 1;
                }
                return 0;
            }
        }

    }
}

