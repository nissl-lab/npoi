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
namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXmlFormats.Wordprocessing;

    /**
     * Experimental class to offer rudimentary Read-only Processing of
     * of StructuredDocumentTags/ContentControl that can appear
     * in a table row as if a table cell.
     * <p/>
     * These can contain one or more cells or other SDTs within them.
     * <p/>
     * WARNING - APIs expected to change rapidly
     */
    public class XWPFSDTCell : AbstractXWPFSDT, ICell
    {
        private XWPFSDTContentCell cellContent;

        public XWPFSDTCell(CT_SdtCell sdtCell, XWPFTableRow xwpfTableRow, IBody part)
            : base(sdtCell.sdtPr, part)
        {
            //base(sdtCell.SdtPr, part);
            cellContent = new XWPFSDTContentCell(sdtCell.sdtContent, xwpfTableRow, part);
        }


        public override ISDTContent Content
        {
            get
            {
                return cellContent;
            }
        }

    }
}
