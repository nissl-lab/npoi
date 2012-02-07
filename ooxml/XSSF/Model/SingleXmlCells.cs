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
using java.util.List;
using java.util.Vector;

using org.apache.poi.POIXMLDocumentPart;
using org.apache.poi.Openxml4j.opc.PackagePart;
using org.apache.poi.Openxml4j.opc.PackageRelationship;
using NPOI.XSSF.usermodel.XSSFSheet;
using NPOI.XSSF.usermodel.helpers.XSSFSingleXmlCell;
using org.apache.xmlbeans.XmlException;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTSingleXmlCell;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTSingleXmlCells;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.SingleXmlCellsDocument;


/**
 * 
 * This class : the Single Cell Tables Part (Open Office XML Part 4:
 * chapter 3.5.2)
 * 
 *
 * @author Roberto Manicardi
 */
public class SingleXmlCells : POIXMLDocumentPart {
	
	
	private CTSingleXmlCells SingleXMLCells;

	public SingleXmlCells() {
		base();
		SingleXMLCells = CTSingleXmlCells.Factory.newInstance();

	}

	public SingleXmlCells(PackagePart part, PackageRelationship rel)
		  {
		base(part, rel);
		ReadFrom(part.GetStream());
	}

	public void ReadFrom(Stream is)  {
		try {
			SingleXmlCellsDocument doc = SingleXmlCellsDocument.Factory.Parse(is);
			SingleXMLCells = doc.GetSingleXmlCells();
		} catch (XmlException e) {
			throw new IOException(e.GetLocalizedMessage());
		}
	}
	
	public XSSFSheet GetXSSFSheet(){
		return (XSSFSheet) GetParent();
	}

	protected void WriteTo(Stream out)  {
		SingleXmlCellsDocument doc = SingleXmlCellsDocument.Factory.newInstance();
		doc.SetSingleXmlCells(SingleXMLCells);
		doc.save(out, DEFAULT_XML_OPTIONS);
	}

	
	protected void Commit()  {
		PackagePart part = GetPackagePart();
		Stream out = part.GetStream();
		WriteTo(out);
		out.Close();
	}
	
	public CTSingleXmlCells GetCTSingleXMLCells(){
		return SingleXMLCells;
	}
	
	/**
	 * 
	 * @return all the SimpleXmlCell Contained in this SingleXmlCells element
	 */
	public List<XSSFSingleXmlCell> GetAllSimpleXmlCell(){
		List<XSSFSingleXmlCell> list = new Vector<XSSFSingleXmlCell>();
		
		for(CTSingleXmlCell SingleXmlCell: SingleXMLCells.GetSingleXmlCellList()){			
			list.Add(new XSSFSingleXmlCell(SingleXmlCell,this));
		}		
		return list;
	}
}






