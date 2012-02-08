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

namespace NPOI.xssf.usermodel;




using NPOI.POIXMLDocumentPart;
using NPOI.util.Internal;
using NPOI.xssf.Model.MapInfo;
using NPOI.xssf.Model.SingleXmlCells;
using NPOI.xssf.usermodel.helpers.XSSFSingleXmlCell;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTMap;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTSchema;
using org.w3c.dom.Node;

/**
 * This class : the Map element (Open Office XML Part 4:
 * chapter 3.16.2)
 * <p/>
 * This element Contains all of the properties related to the XML map,
 * and the behaviors expected during data refresh operations.
 *
 * @author Roberto Manicardi
 */


public class XSSFMap {

    private CTMap ctMap;

    private MapInfo mapInfo;


    public XSSFMap(CTMap ctMap, MapInfo mapInfo) {
        this.ctMap = ctMap;
        this.mapInfo = mapInfo;
    }


    
    public CTMap GetCtMap() {
        return ctMap;
    }


    
    public CTSchema GetCTSchema() {
        String schemaId = ctMap.GetSchemaID();
        return mapInfo.GetCTSchemaById(schemaId);
    }

    public Node GetSchema() {
        Node xmlSchema = null;

        CTSchema schema = GetCTSchema();
        xmlSchema = schema.GetDomNode().GetFirstChild();

        return xmlSchema;
    }

    /**
     * @return the list of Single Xml Cells that provide a map rule to this mapping.
     */
    public List<XSSFSingleXmlCell> GetRelatedSingleXMLCell() {
        List<XSSFSingleXmlCell> relatedSimpleXmlCells = new Vector<XSSFSingleXmlCell>();

        int sheetNumber = mapInfo.GetWorkbook().GetNumberOfSheets();
        for (int i = 0; i < sheetNumber; i++) {
            XSSFSheet sheet = mapInfo.GetWorkbook().GetSheetAt(i);
            foreach (POIXMLDocumentPart p in sheet.GetRelations()) {
                if (p is SingleXmlCells) {
                    SingleXmlCells SingleXMLCells = (SingleXmlCells) p;
                    foreach (XSSFSingleXmlCell cell in SingleXMLCells.GetAllSimpleXmlCell()) {
                        if (cell.GetMapId() == ctMap.GetID()) {
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
    public List<XSSFTable> GetRelatedTables() {

        List<XSSFTable> tables = new Vector<XSSFTable>();
        int sheetNumber = mapInfo.GetWorkbook().GetNumberOfSheets();

        for (int i = 0; i < sheetNumber; i++) {
            XSSFSheet sheet = mapInfo.GetWorkbook().GetSheetAt(i);
            foreach (POIXMLDocumentPart p in sheet.GetRelations()) {
                if (p.GetPackageRelationship().GetRelationshipType().Equals(XSSFRelation.TABLE.GetRelation())) {
                    XSSFTable table = (XSSFTable) p;
                    if (table.mapsTo(ctMap.GetID())) {
                        tables.Add(table);
                    }
                }
            }
        }

        return tables;
    }
}


