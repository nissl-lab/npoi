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
using NPOI.Openxml4j.opc.PackagePart;
using NPOI.Openxml4j.opc.PackageRelationship;
using NPOI.xssf.util.EvilUnclosedBRFixingInputStream;
using org.apache.xmlbeans.XmlException;
using org.apache.xmlbeans.XmlOptions;
using org.apache.xmlbeans.XmlObject;
using org.apache.xmlbeans.XmlCursor;
using org.w3c.dom.Node;
using schemasMicrosoftComOfficeOffice.*;

using javax.xml.namespace.QName;






using schemasMicrosoftComVml.*;
using schemasMicrosoftComVml.STTrueFalse;
using schemasMicrosoftComOfficeExcel.CTClientData;
using schemasMicrosoftComOfficeExcel.STObjectType;

/**
 * Represents a SpreadsheetML VML Drawing.
 *
 * <p>
 * In Excel 2007 VML Drawings are used to describe properties of cell comments,
 * although the spec says that VML is deprecated:
 * </p>
 * <p>
 * The VML format is a legacy format originally introduced with Office 2000 and is included and fully defined
 * in this Standard for backwards compatibility reasons. The DrawingML format is a newer and richer format
 * Created with the goal of eventually replacing any uses of VML in the Office Open XML formats. VML should be
 * considered a deprecated format included in Office Open XML for legacy reasons only and new applications that
 * need a file format for Drawings are strongly encouraged to use preferentially DrawingML
 * </p>
 * 
 * <p>
 * Warning - Excel is known to Put invalid XML into these files!
 *  For example, &gt;br&lt; without being closed or escaped crops up.
 * </p>
 *
 * See 6.4 VML - SpreadsheetML Drawing in Office Open XML Part 4 - Markup Language Reference.pdf
 *
 * @author Yegor Kozlov
 */
public class XSSFVMLDrawing : POIXMLDocumentPart {
    private static QName QNAME_SHAPE_LAYOUT = new QName("urn:schemas-microsoft-com:office:office", "shapelayout");
    private static QName QNAME_SHAPE_TYPE = new QName("urn:schemas-microsoft-com:vml", "shapetype");
    private static QName QNAME_SHAPE = new QName("urn:schemas-microsoft-com:vml", "shape");

    /**
     * regexp to parse shape ids, in VML they have weird form of id="_x0000_s1026"
     */
    private static Pattern ptrn_shapeId = Pattern.compile("_x0000_s(\\d+)");

    private List<QName> _qnames = new List<QName>();
    private List<XmlObject> _items = new List<XmlObject>();
    private String _shapeTypeId;
    private int _shapeId = 1024;

    /**
     * Create a new SpreadsheetML Drawing
     *
     * @see XSSFSheet#CreateDrawingPatriarch()
     */
    protected XSSFVMLDrawing() {
        base();
        newDrawing();
    }

    /**
     * Construct a SpreadsheetML Drawing from a namespace part
     *
     * @param part the namespace part holding the Drawing data,
     * the content type must be <code>application/vnd.Openxmlformats-officedocument.Drawing+xml</code>
     * @param rel  the namespace relationship holding this Drawing,
     * the relationship type must be http://schemas.Openxmlformats.org/officeDocument/2006/relationships/drawing
     */
    protected XSSFVMLDrawing(PackagePart part, PackageRelationship rel) , XmlException {
        base(part, rel);
        Read(getPackagePart().GetInputStream());
    }


    protected void Read(InputStream is) , XmlException {
        XmlObject root = XmlObject.Factory.Parse(
              new EvilUnclosedBRFixingInputStream(is)
        );

        _qnames = new List<QName>();
        _items = new List<XmlObject>();
        foreach(XmlObject obj in root.selectPath("$this/xml/*")) {
            Node nd = obj.GetDomNode();
            QName qname = new QName(nd.GetNamespaceURI(), nd.GetLocalName());
            if (qname.Equals(QNAME_SHAPE_LAYOUT)) {
                _items.Add(CTShapeLayout.Factory.Parse(obj.xmlText()));
            } else if (qname.Equals(QNAME_SHAPE_TYPE)) {
                CTShapetype st = CTShapetype.Factory.Parse(obj.xmlText());
                _items.Add(st);
                _shapeTypeId = st.GetId();
            } else if (qname.Equals(QNAME_SHAPE)) {
                CTShape shape = CTShape.Factory.Parse(obj.xmlText());
                String id = shape.GetId();
                if(id != null) {
                    Matcher m = ptrn_shapeId.matcher(id);
                    if(m.Find()) _shapeId = Math.max(_shapeId, Int32.ParseInt(m.group(1)));
                }
                _items.Add(shape);
            } else {
                _items.Add(XmlObject.Factory.Parse(obj.xmlText()));
            }
            _qnames.Add(qname);
        }
    }

    protected List<XmlObject> GetItems(){
        return _items;
    }

    protected void Write(OutputStream out)  {
        XmlObject rootObject = XmlObject.Factory.newInstance();
        XmlCursor rootCursor = rootObject.newCursor();
        rootCursor.ToNextToken();
        rootCursor.beginElement("xml");

        for(int i=0; i < _items.Count; i++){
            XmlCursor xc = _items.Get(i).newCursor();
            rootCursor.beginElement(_qnames.Get(i));
            while(xc.ToNextToken() == XmlCursor.TokenType.ATTR) {
                Node anode = xc.GetDomNode();
                rootCursor.insertAttributeWithValue(anode.GetLocalName(), anode.GetNamespaceURI(), anode.GetNodeValue());
            }
            xc.ToStartDoc();
            xc.copyXmlContents(rootCursor);
            rootCursor.ToNextToken();
            xc.dispose();
        }
        rootCursor.dispose();

        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        xmlOptions.SetSavePrettyPrint();
        Dictionary<String, String> map = new Dictionary<String, String>();
        map.Put("urn:schemas-microsoft-com:vml", "v");
        map.Put("urn:schemas-microsoft-com:office:office", "o");
        map.Put("urn:schemas-microsoft-com:office:excel", "x");
        xmlOptions.SetSaveSuggestedPrefixes(map);

        rootObject.save(out, xmlOptions);
    }

    
    protected void Commit()  {
        PackagePart part = GetPackagePart();
        OutputStream out = part.GetOutputStream();
        Write(out);
        out.Close();
    }

    /**
     * Initialize a new Speadsheet VML Drawing
     */
    private void newDrawing(){
        CTShapeLayout layout = CTShapeLayout.Factory.newInstance();
        layout.SetExt(STExt.EDIT);
        CTIdMap idmap = layout.AddNewIdmap();
        idmap.SetExt(STExt.EDIT);
        idmap.SetData("1");
        _items.Add(layout);
        _qnames.Add(QNAME_SHAPE_LAYOUT);

        CTShapetype shapetype = CTShapetype.Factory.newInstance();
        _shapeTypeId = "_xssf_cell_comment";
        shapetype.SetId(_shapeTypeId);
        shapetype.SetCoordsize("21600,21600");
        shapetype.SetSpt(202);
        shapetype.SetPath2("m,l,21600r21600,l21600,xe");
        shapetype.AddNewStroke().SetJoinstyle(STStrokeJoinStyle.MITER);
        CTPath path = shapetype.AddNewPath();
        path.SetGradientshapeok(STTrueFalse.T);
        path.SetConnecttype(STConnectType.RECT);
        _items.Add(shapetype);
        _qnames.Add(QNAME_SHAPE_TYPE);
    }

    protected CTShape newCommentShape(){
        CTShape shape = CTShape.Factory.newInstance();
        shape.SetId("_x0000_s" + (++_shapeId));
        shape.SetType("#" + _shapeTypeId);
        shape.SetStyle("position:absolute; visibility:hidden");
        shape.SetFillcolor("#ffffe1");
        shape.SetInsetmode(STInsetMode.AUTO);
        shape.AddNewFill().SetColor("#ffffe1");
        CTShadow shadow = shape.AddNewShadow();
        shadow.SetOn(STTrueFalse.T);
        shadow.SetColor("black");
        shadow.SetObscured(STTrueFalse.T);
        shape.AddNewPath().SetConnecttype(STConnectType.NONE);
        shape.AddNewTextbox().SetStyle("mso-direction-alt:auto");
        CTClientData cldata = shape.AddNewClientData();
        cldata.SetObjectType(STObjectType.NOTE);
        cldata.AddNewMoveWithCells();
        cldata.AddNewSizeWithCells();
        cldata.AddNewAnchor().SetStringValue("1, 15, 0, 2, 3, 15, 3, 16");
        cldata.AddNewAutoFill().SetStringValue("False");
        cldata.AddNewRow().SetBigintValue(new Bigint("0"));
        cldata.AddNewColumn().SetBigintValue(new Bigint("0"));
        _items.Add(shape);
        _qnames.Add(QNAME_SHAPE);
        return shape;
    }

    /**
     * Find a shape with ClientData of type "NOTE" and the specified row and column
     *
     * @return the comment shape or <code>null</code>
     */
    protected CTShape FindCommentShape(int row, int col){
        foreach(XmlObject itm in _items){
            if(itm is CTShape){
                CTShape sh = (CTShape)itm;
                if(sh.sizeOfClientDataArray() > 0){
                    CTClientData cldata = sh.GetClientDataArray(0);
                    if(cldata.GetObjectType() == STObjectType.NOTE){
                        int crow = cldata.GetRowArray(0).intValue();
                        int ccol = cldata.GetColumnArray(0).intValue();
                        if(crow == row && ccol == col) {
                            return sh;
                        }
                    }
                }
            }
        }
        return null;
    }

    protected bool RemoveCommentShape(int row, int col){
        CTShape shape = FindCommentShape(row, col);
        return shape != null && _items.Remove(shape);
    }
}

