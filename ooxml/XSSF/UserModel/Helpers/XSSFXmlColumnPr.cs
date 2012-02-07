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

namespace NPOI.XSSF.usermodel.helpers;

using NPOI.XSSF.usermodel.XSSFTable;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTXmlColumnPr;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STXmlDataType.Enum;


/**
 * 
 * This class is a wrapper around the CTXmlColumnPr (Open Office XML Part 4:
 * chapter 3.5.1.7)
 * 
 *
 * @author Roberto Manicardi
 */
public class XSSFXmlColumnPr {
	
	private XSSFTable table;
	private CTTableColumn ctTableColumn;
	private CTXmlColumnPr ctXmlColumnPr;
	
	public XSSFXmlColumnPr(XSSFTable table ,CTTableColumn ctTableColum,CTXmlColumnPr ctXmlColumnPr){
		this.table = table;
		this.ctTableColumn = ctTableColum;
		this.ctXmlColumnPr = ctXmlColumnPr;
	}
	
	public long GetMapId(){
		return ctXmlColumnPr.GetMapId();
	}
	
	public String GetXPath(){
		return ctXmlColumnPr.GetXpath();
	}
	/**
	 * (see Open Office XML Part 4: chapter 3.5.1.3)
	 * @return An integer representing the unique identifier of this column. 
	 */
	public long GetId(){
		return ctTableColumn.GetId();
	}
	
	
	/**
	 * If the XPath is, for example, /Node1/Node2/Node3 and /Node1/Node2 is the common XPath for the table, the local XPath is /Node3
	 * 	
	 * @return the local XPath 
	 */
	public String GetLocalXPath(){
		String localXPath = "";
		int numberOfCommonXPathAxis = table.GetCommonXpath().split("/").Length-1;
		
		String[] xPathTokens = ctXmlColumnPr.GetXpath().split("/");
		for(int i=numberOfCommonXPathAxis; i<xPathTokens.Length;i++){
			localXPath += "/" +xPathTokens[i];
		}
		return localXPath;
	}

	public Enum GetXmlDataType() {
		
		return ctXmlColumnPr.GetXmlDataType();
	}
	
	
	
	

}


