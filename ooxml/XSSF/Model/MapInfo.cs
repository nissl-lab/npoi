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

namespace NPOI.XSSF.model;

using java.io.IOException;
using java.io.Stream;
using java.io.Stream;
using java.util.Collection;
using java.util.HashMap;
using java.util.Map;


using org.apache.poi.POIXMLDocumentPart;
using org.apache.poi.Openxml4j.opc.PackagePart;
using org.apache.poi.Openxml4j.opc.PackageRelationship;
using NPOI.XSSF.usermodel.XSSFMap;
using NPOI.XSSF.usermodel.XSSFWorkbook;
using org.apache.xmlbeans.XmlException;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTMap;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTMapInfo;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTSchema;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.MapInfoDocument;

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

public class MapInfo : POIXMLDocumentPart {

	private CTMapInfo mapInfo;
	
	private Dictionary<int, XSSFMap> maps ;

	public MapInfo() {
		base();
		mapInfo = CTMapInfo.Factory.newInstance();

	}

	public MapInfo(PackagePart part, PackageRelationship rel)
		  {
		base(part, rel);
		ReadFrom(part.GetStream());
	}

	public void ReadFrom(Stream is)  {
		try {
			MapInfoDocument doc = MapInfoDocument.Factory.Parse(is);
			mapInfo = doc.GetMapInfo();

            maps= new Dictionary<int, XSSFMap>();
            for(CTMap map :mapInfo.GetMapList()){
                maps.Put((int)map.GetID(), new XSSFMap(map,this));
            }

		} catch (XmlException e) {
			throw new IOException(e.GetLocalizedMessage());
		}
	}
	
	/**
     * Returns the parent XSSFWorkbook
     *
     * @return the parent XSSFWorkbook
     */
    public XSSFWorkbook GetWorkbook() {
        return (XSSFWorkbook)GetParent();
    }
	
	/**
	 * 
	 * @return the internal data object
	 */
	public CTMapInfo GetCTMapInfo(){
		return mapInfo;
		
	}

	/**
	 * Gets the
	 * @param schemaId the schema ID
	 * @return CTSchema by it's ID
	 */
	public CTSchema GetCTSchemaById(String schemaId){
		CTSchema xmlSchema = null;

		for(CTSchema schema: mapInfo.GetSchemaList()){
			if(schema.GetID().Equals(schemaId)){
				xmlSchema = schema;
				break;
			}
		}
		return xmlSchema;
	}
	
	
	public XSSFMap GetXSSFMapById(int id){
		return maps.Get(id);
	}
	
	public XSSFMap GetXSSFMapByName(String name){
		
		XSSFMap matchedMap = null;
		
		for(XSSFMap map :maps.values()){
			if(map.GetCtMap().GetName()!=null && map.GetCtMap().GetName().Equals(name)){
				matchedMap = map;
			}
		}		
		
		return matchedMap;
	}
	
	/**
	 * 
	 * @return all the mappings configured in this document
	 */
	public Collection<XSSFMap> GetAllXSSFMaps(){
		return maps.values();
	}

	protected void WriteTo(Stream out)  {
		MapInfoDocument doc = MapInfoDocument.Factory.newInstance();
		doc.SetMapInfo(mapInfo);
		doc.save(out, DEFAULT_XML_OPTIONS);
	}

	
	protected void Commit()  {
		PackagePart part = GetPackagePart();
		Stream out = part.GetStream();
		WriteTo(out);
		out.Close();
	}

}






