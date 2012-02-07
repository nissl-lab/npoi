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

using NPOI.SS.util.CellReference;
using NPOI.XSSF.model.SingleXmlCells;
using NPOI.XSSF.usermodel.XSSFCell;
using NPOI.XSSF.usermodel.XSSFRow;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTSingleXmlCell;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTXmlCellPr;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTXmlPr;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.STXmlDataType.Enum;

/**
 * 
 * This class is a wrapper around the CTSingleXmlCell  (Open Office XML Part 4:
 * chapter 3.5.2.1) 
 * 

 * 
 * @author Roberto Manicardi
 *
 */
public class XSSFSingleXmlCell {
	
	private CTSingleXmlCell SingleXmlCell;
	private SingleXmlCells parent;
	
	
	public XSSFSingleXmlCell(CTSingleXmlCell SingleXmlCell, SingleXmlCells parent){
		this.SingleXmlCell = SingleXmlCell;
		this.parent = parent;
	}
	
	/**
	 * Gets the XSSFCell referenced by the R attribute or Creates a new one if cell doesn't exists
	 * @return the referenced XSSFCell, null if the cell reference is invalid
	 */
	public XSSFCell GetReferencedCell(){
		XSSFCell cell = null;
		
		
		CellReference cellReference =  new CellReference(SingleXmlCell.GetR()); 
		
		XSSFRow row = parent.GetXSSFSheet().GetRow(cellReference.getRow());
		if(row==null){
			row = parent.GetXSSFSheet().CreateRow(cellReference.getRow());
		}
		
		cell = row.GetCell(cellReference.GetCol());  
		if(cell==null){
			cell = row.CreateCell(cellReference.GetCol());
		}
		
		
		return cell;
	}
	
	public String GetXpath(){
		CTXmlCellPr xmlCellPr = SingleXmlCell.GetXmlCellPr();
		CTXmlPr xmlPr = xmlCellPr.GetXmlPr();
		String xpath = xmlPr.GetXpath();
		return xpath;
	}
	
	public long GetMapId(){
		return SingleXmlCell.GetXmlCellPr().getXmlPr().getMapId();
	}

	public Enum GetXmlDataType() {
		CTXmlCellPr xmlCellPr = SingleXmlCell.GetXmlCellPr();
		CTXmlPr xmlPr = xmlCellPr.GetXmlPr();
		return xmlPr.GetXmlDataType();
	}

}


