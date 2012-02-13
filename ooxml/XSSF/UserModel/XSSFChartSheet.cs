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









using javax.xml.namespace.QName;

using NPOI.POIXMLException;
using NPOI.Openxml4j.opc.PackagePart;
using NPOI.Openxml4j.opc.PackageRelationship;
using org.apache.xmlbeans.XmlException;
using org.apache.xmlbeans.XmlOptions;
using org.Openxmlformats.schemas.officeDocument.x2006.relationships.STRelationshipId;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTChartsheet;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTDrawing;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTLegacyDrawing;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.ChartsheetDocument;

/**
 * High level representation of Sheet Parts that are of type 'chartsheet'.
 * <p>
 *  Chart sheet is a special kind of Sheet that Contains only chart and no data.
 * </p>
 *
 * @author Yegor Kozlov
 */
public class XSSFChartSheet : XSSFSheet  {

    private static byte[] BLANK_WORKSHEET = blankWorksheet();

    protected CTChartsheet chartsheet;

    protected XSSFChartSheet(PackagePart part, PackageRelationship rel) {
        base(part, rel);
    }

    protected void Read(InputStream is)  {
        //Initialize the supeclass with a blank worksheet
        super.Read(new MemoryStream(BLANK_WORKSHEET));

        try {
            chartsheet = ChartsheetDocument.Factory.Parse(is).GetChartsheet();
        } catch (XmlException e){
            throw new POIXMLException(e);
        }
    }

    /**
     * Provide access to the CTChartsheet bean holding this sheet's data
     *
     * @return the CTChartsheet bean holding this sheet's data
     */
    public CTChartsheet GetCTChartsheet() {
        return chartsheet;
    }

    
    protected CTDrawing GetCTDrawing() {
       return chartsheet.GetDrawing();
    }
    
    
    protected CTLegacyDrawing GetCTLegacyDrawing() {
       return chartsheet.GetLegacyDrawing();
    }
    
    
    protected void Write(OutputStream out)  {
        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        xmlOptions.SetSaveSyntheticDocumentElement(
                new QName(CTChartsheet.type.GetName().GetNamespaceURI(), "chartsheet"));
        Dictionary<String, String> map = new Dictionary<String, String>();
        map.Put(STRelationshipId.type.GetName().GetNamespaceURI(), "r");
        xmlOptions.SetSaveSuggestedPrefixes(map);

        chartsheet.save(out, xmlOptions);

    }

    private static byte[] blankWorksheet(){
        MemoryStream out = new MemoryStream();
        try {
            new XSSFSheet().Write(out);
        } catch (IOException e){
            throw new RuntimeException(e);
        }
        return out.ToArray();
    }
}


