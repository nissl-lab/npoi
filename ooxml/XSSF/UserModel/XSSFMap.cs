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
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.UserModel.Helpers;
using System.Collections.Generic;
using NPOI.XSSF.Model;
using NPOI.SS.UserModel;

namespace NPOI.XSSF.UserModel
{
    /**
     * This class : the Map element (Open Office XML Part 4:
     * chapter 3.16.2)
     * 
     * This element Contains all of the properties related to the XML map,
     * and the behaviors expected during data refresh operations.
     *
     * @author Roberto Manicardi
     */


    public class XSSFMap
    {

        private CT_Map ctMap;

        private MapInfo mapInfo;


        public XSSFMap(CT_Map ctMap, MapInfo mapInfo)
        {
            this.ctMap = ctMap;
            this.mapInfo = mapInfo;
        }



        public CT_Map GetCTMap()
        {
            return ctMap;
        }



        public CT_Schema GetCTSchema()
        {
            String schemaId = ctMap.SchemaID;
            return mapInfo.GetCTSchemaById(schemaId);
        }

        public string GetSchema()
        {
            CT_Schema ctSchema = GetCTSchema();
            return ctSchema.InnerXml;
        }

        /**
         * @return the list of Single Xml Cells that provide a map rule to this mapping.
         */
        public List<XSSFSingleXmlCell> GetRelatedSingleXMLCell()
        {
            List<XSSFSingleXmlCell> relatedSimpleXmlCells = new List<XSSFSingleXmlCell>();

            int sheetNumber = mapInfo.Workbook.NumberOfSheets;
            for (int i = 0; i < sheetNumber; i++)
            {
                XSSFSheet sheet = (XSSFSheet)mapInfo.Workbook.GetSheetAt(i);
                foreach (POIXMLDocumentPart p in sheet.GetRelations())
                {
                    if (p is SingleXmlCells)
                    {
                        SingleXmlCells SingleXMLCells = (SingleXmlCells)p;
                        foreach (XSSFSingleXmlCell cell in SingleXMLCells.GetAllSimpleXmlCell())
                        {
                            if (cell.GetMapId() == ctMap.ID)
                            {
                                relatedSimpleXmlCells.Add(cell);
                            }
                        }
                    }
                }
            }
            return relatedSimpleXmlCells;
        }

        /**
         * @return the list of all Tables that provide a map rule to this mapping
         */
        public List<XSSFTable> GetRelatedTables()
        {

            List<XSSFTable> tables = new List<XSSFTable>();
            int sheetNumber = mapInfo.Workbook.NumberOfSheets;

            foreach (ISheet sheet in mapInfo.Workbook)
            {
                foreach (POIXMLDocumentPart.RelationPart rp in ((XSSFSheet)sheet).RelationParts)
                {
                    if (rp.Relationship.RelationshipType.Equals(XSSFRelation.TABLE.Relation))
                    {
                        XSSFTable table = rp.DocumentPart as XSSFTable;
                        if (table.MapsTo(ctMap.ID))
                        {
                            tables.Add(table);
                        }
                    }
                }
            }
            return tables;
        }
    }
}


