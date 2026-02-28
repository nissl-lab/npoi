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
using System;
using NPOI.SS.Util;
using NPOI.XSSF.Model;
using NPOI.SS.UserModel;
namespace NPOI.XSSF.UserModel.Helpers
{

    /**
     * 
     * This class is a wrapper around the CTSingleXmlCell  (Open Office XML Part 4:
     * chapter 3.5.2.1) 
     * 

     * 
     * @author Roberto Manicardi
     *
     */
    public class XSSFSingleXmlCell
    {

        private CT_SingleXmlCell SingleXmlCell;
        private SingleXmlCells parent;


        public XSSFSingleXmlCell(CT_SingleXmlCell SingleXmlCell, SingleXmlCells parent)
        {
            this.SingleXmlCell = SingleXmlCell;
            this.parent = parent;
        }

        /**
         * Gets the XSSFCell referenced by the R attribute or Creates a new one if cell doesn't exists
         * @return the referenced XSSFCell, null if the cell reference is invalid
         */
        public ICell GetReferencedCell()
        {
            ICell cell = null;


            CellReference cellReference = new CellReference(SingleXmlCell.r);

            IRow row = parent.GetXSSFSheet().GetRow(cellReference.Row);
            if (row == null)
            {
                row = parent.GetXSSFSheet().CreateRow(cellReference.Row);
            }

            cell = row.GetCell(cellReference.Col);
            if (cell == null)
            {
                cell = row.CreateCell(cellReference.Col);
            }


            return cell;
        }

        public String GetXpath()
        {
            CT_XmlCellPr xmlCellPr = SingleXmlCell.xmlCellPr;
            CT_XmlPr xmlPr = xmlCellPr.xmlPr;
            String xpath = xmlPr.xpath;
            return xpath;
        }

        public long GetMapId()
        {
            return SingleXmlCell.xmlCellPr.xmlPr.mapId;
        }

        public ST_XmlDataType GetXmlDataType()
        {
            CT_XmlCellPr xmlCellPr = SingleXmlCell.xmlCellPr;
            CT_XmlPr xmlPr = xmlCellPr.xmlPr;
            return xmlPr.xmlDataType;
        }

    }


}