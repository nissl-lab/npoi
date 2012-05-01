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

using System.Collections.Generic;
using System;
using NPOI.XSSF.Util;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using System.Text.RegularExpressions;
using NPOI.OpenXmlFormats.Vml;
using System.Xml.Serialization;
using System.Xml;
using System.Collections;
namespace NPOI.XSSF.UserModel
{


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
    public class XSSFVMLDrawing : POIXMLDocumentPart
    {

        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] 
        {
            new XmlQualifiedName("shapelayout", "urn:schemas-microsoft-com:office:office"),
            new XmlQualifiedName("shapetype", "urn:schemas-microsoft-com:vml"),
            new XmlQualifiedName("shape", "urn:schemas-microsoft-com:vml"),
        });

        /**
         * regexp to parse shape ids, in VML they have weird form of id="_x0000_s1026"
         */
        private static Regex ptrn_shapeId = new Regex("_x0000_s(\\d+)");

        private ArrayList _items = new ArrayList();
        private String _shapeTypeId;
        private int _shapeId = 1024;

        /**
         * Create a new SpreadsheetML Drawing
         *
         * @see XSSFSheet#CreateDrawingPatriarch()
         */
        public XSSFVMLDrawing()
            : base()
        {

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
        protected XSSFVMLDrawing(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

            Read(GetPackagePart().GetInputStream());
        }


        protected void Read(Stream is1)
        {
            XmlDocument doc = new XmlDocument();
             doc.Load(
                  new EvilUnclosedBRFixingInputStream(is1)
            );

             XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
             nsmgr.AddNamespace("shapelayout", "urn:schemas-microsoft-com:office:office");
             nsmgr.AddNamespace("shapetype", "urn:schemas-microsoft-com:vml");
             nsmgr.AddNamespace("shape", "urn:schemas-microsoft-com:vml");
             _items = new ArrayList();
            XmlNodeList nodes=doc.SelectNodes("*",nsmgr);
            foreach (XmlNode nd in nodes)
            {
            }
        }

        protected ArrayList GetItems()
        {
            return _items;
        }

        protected void Write(Stream out1)
        {
            throw new NotImplementedException();
            //XmlObject rootObject = XmlObject();
            //XmlCursor rootCursor = rootObject.newCursor();
            //rootCursor.ToNextToken();
            //rootCursor.beginElement("xml");

            //for (int i = 0; i < _items.Count; i++)
            //{
            //    XmlCursor xc = _items[i].newCursor();
            //    rootCursor.beginElement(_qnames[i]);
            //    while (xc.ToNextToken() == XmlCursor.TokenType.ATTR)
            //    {
            //        Node anode = xc.GetDomNode();
            //        rootCursor.insertAttributeWithValue(anode.GetLocalName(), anode.GetNamespaceURI(), anode.GetNodeValue());
            //    }
            //    xc.ToStartDoc();
            //    xc.copyXmlContents(rootCursor);
            //    rootCursor.ToNextToken();
            //    xc.dispose();
            //}
            //rootCursor.dispose();


            //Dictionary<String, String> map = new Dictionary<String, String>();
            //map["urn:schemas-microsoft-com:vml"] = "v";
            //map["urn:schemas-microsoft-com:office:office"] = "o";
            //map["urn:schemas-microsoft-com:office:excel"] = "x";

            //rootObject.save(out1);
        }


        protected override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            Write(out1);
            out1.Close();
        }

        /**
         * Initialize a new Speadsheet VML Drawing
         */
        private void newDrawing()
        {
            CT_ShapeLayout layout = new CT_ShapeLayout();
            layout.ext = (ST_Ext.edit);
            CT_IdMap idmap = layout.AddNewIdmap();
            idmap.ext = (ST_Ext.edit);
            idmap.data = ("1");
            _items.Add(layout);

            //CT_Shapetype shapetype = new CT_Shapetype();
            //_shapeTypeId = "_xssf_cell_comment";
            //shapetype.id= _shapeTypeId; 
            //shapetype.SetCoordsize("21600,21600");
            //shapetype.SetSpt(202);
            //shapetype.path1 = ("m,l,21600r21600,l21600,xe");
            //shapetype.AddNewStroke().SetJoinstyle(ST_StrokeJoinStyle.miter);
            //CT_Path path = shapetype.AddNewPath();
            //path.gradientshapeok = ST_TrueFalse.t;
            //path.SetConnecttype(ST_ConnectType.rect);
            //_items.Add(shapetype);
        }

        internal CT_Shape newCommentShape()
        {
            CT_Shape shape = new CT_Shape();

            //shape.SetId("_x0000_s" + (++_shapeId));
            //shape.SetType("#" + _shapeTypeId);
            //shape.SetStyle("position:absolute; visibility:hidden");
            //shape.SetFillcolor("#ffffe1");
            //shape.SetInsetmode(ST_InsetMode.auto);
            //shape.AddNewFill().color=("#ffffe1");
            //CT_Shadow shadow = shape.AddNewShadow();
            //shadow.on= ST_TrueFalse.t;
            //shadow.color = "black";
            //shadow.obscured = ST_TrueFalse.t;
            //shape.AddNewPath().SetConnecttype(ST_ConnectType.none);
            //shape.AddNewTextbox().SetStyle("mso-direction-alt:auto");
            //CT_ClientData cldata = shape.AddNewClientData();
            //cldata.ObjectType=ST_ObjectType.Note;
            //cldata.AddNewMoveWithCells();
            //cldata.AddNewSizeWithCells();
            //cldata.AddNewAnchor().SetStringValue("1, 15, 0, 2, 3, 15, 3, 16");
            //cldata.AddNewAutoFill().SetStringValue("False");
            //cldata.AddNewRow().SetBigintValue(new Bigint("0"));
            //cldata.AddNewColumn().SetBigintValue(new Bigint("0"));
            //_items.Add(shape);

            return shape;
        }

        /**
         * Find a shape with ClientData of type "NOTE" and the specified row and column
         *
         * @return the comment shape or <code>null</code>
         */
        internal CT_Shape FindCommentShape(int row, int col)
        {
            //foreach (XmlObject itm in _items)
            //{
            //    if (itm is CT_Shape)
            //    {
            //        CT_Shape sh = (CT_Shape)itm;
            //        if (sh.sizeOfClientDataArray() > 0)
            //        {
            //            CT_ClientData cldata = sh.GetClientDataArray(0);
            //            if (cldata.GetObjectType() == ST_ObjectType.NOTE)
            //            {
            //                int crow = cldata.GetRowArray(0).intValue();
            //                int ccol = cldata.GetColumnArray(0).intValue();
            //                if (crow == row && ccol == col)
            //                {
            //                    return sh;
            //                }
            //            }
            //        }
            //    }
            //}
            return null;
        }

        internal bool RemoveCommentShape(int row, int col)
        {
            CT_Shape shape = FindCommentShape(row, col);
            return shape != null; //&& _items.Remove(shape);
        }
    }
}

