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

using System.Xml;
using System.Collections.Generic;
using System.IO;
using NPOI.OpenXmlFormats;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using NPOI.XSSF.UserModel;
using NPOI.Util;
namespace NPOI.XSSF.Model
{


    /**
     * 
     * This class : the Custom XML Mapping Part (Open Office XML Part 1:
     * chapter 12.3.6)
     * 
     * An instance of this part type Contains a schema for an XML file, and
     * information on the behavior that is used when allowing this custom XML schema
     * to be mapped into the spreadsheet.
     * 
     * @author Roberto Manicardi
     */

    public class MapInfo : POIXMLDocumentPart
    {

        private CT_MapInfo mapInfo;

        private Dictionary<int, XSSFMap> maps;

        public MapInfo()
            : base()
        {

            mapInfo = new CT_MapInfo();

        }
        XmlDocument xml = null;
        internal MapInfo(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            xml = ConvertStreamToXml(part.GetInputStream());
            ReadFrom(xml);
        }

        public void ReadFrom(XmlDocument xmldoc)
        {
            try
            {
                MapInfoDocument doc = MapInfoDocument.Parse(xmldoc, NamespaceManager);
                mapInfo = doc.GetMapInfo();

                maps = new Dictionary<int, XSSFMap>();
                foreach (CT_Map map in mapInfo.Map)
                {
                    maps[(int)map.ID] = new XSSFMap(map, this);
                }

            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }

        /**
         * Returns the parent XSSFWorkbook
         *
         * @return the parent XSSFWorkbook
         */
        public XSSFWorkbook Workbook
        {
            get
            {
                return (XSSFWorkbook)GetParent();
            }
        }

        /**
         * 
         * @return the internal data object
         */
        public CT_MapInfo GetCTMapInfo()
        {
            return mapInfo;

        }

        /**
         * Gets the
         * @param schemaId the schema ID
         * @return CTSchema by it's ID
         */
        public NPOI.OpenXmlFormats.Spreadsheet.CT_Schema GetCTSchemaById(String schemaId)
        {
            NPOI.OpenXmlFormats.Spreadsheet.CT_Schema xmlSchema = null;

            foreach (NPOI.OpenXmlFormats.Spreadsheet.CT_Schema schema in mapInfo.Schema)
            {
                if (schema.ID.Equals(schemaId))
                {
                    xmlSchema = schema;
                    break;
                }
            }
            return xmlSchema;
        }


        public XSSFMap GetXSSFMapById(int id)
        {
            return maps[id];
        }

        public XSSFMap GetXSSFMapByName(String name){
		
		XSSFMap matchedMap = null;
		
		foreach(XSSFMap map in maps.Values){
            if (map.GetCTMap().Name != null && map.GetCTMap().Name.Equals(name))
            {
				matchedMap = map;
			}
		}		
		
		return matchedMap;
	}

        /**
         * 
         * @return all the mappings configured in this document
         */
        public List<XSSFMap> GetAllXSSFMaps()
        {
            List<XSSFMap> tmaps = new List<XSSFMap>();
            foreach(XSSFMap map in maps.Values)
            {
                tmaps.Add(map);
            }
            return tmaps;
        }

        protected void WriteTo(Stream out1)
        {
            //MapInfoDocument doc = new MapInfoDocument();
            //doc.SetMapInfo(mapInfo);
            //doc.Save(out1);
            xml.Save(out1);
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }

    }
}



