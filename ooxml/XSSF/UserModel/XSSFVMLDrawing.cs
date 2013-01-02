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
using NPOI.OpenXmlFormats.Vml.Office;
using NPOI.OpenXmlFormats.Vml.Spreadsheet;
using System.Text;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;

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
        private static XmlQualifiedName QNAME_SHAPE_LAYOUT = new XmlQualifiedName("shapelayout", "urn:schemas-microsoft-com:office:office");
        private static XmlQualifiedName QNAME_SHAPE_TYPE = new XmlQualifiedName("shapetype", "urn:schemas-microsoft-com:vml");
        private static XmlQualifiedName QNAME_SHAPE = new XmlQualifiedName("shape", "urn:schemas-microsoft-com:vml");
        
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


        internal void Read(Stream is1)
        {
            XmlDocument doc = new XmlDocument();

            //InflaterInputStream iis = (InflaterInputStream)is1;
            StreamReader sr = new StreamReader(is1);
            string data = sr.ReadToEnd();
            
            //Stream vmlsm = new EvilUnclosedBRFixingInputStream(is1); --TODO:: add later
            
             doc.LoadXml(
                  data.Replace("<br>","")
            );

             XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
             nsmgr.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
             nsmgr.AddNamespace("x", "urn:schemas-microsoft-com:office:excel");
             nsmgr.AddNamespace("v", "urn:schemas-microsoft-com:vml");
             _items = new ArrayList();
            XmlNodeList nodes=doc.SelectNodes("/xml/*",nsmgr);
            foreach (XmlNode nd in nodes)
            {
                string xmltext = nd.OuterXml;
                if (nd.LocalName == QNAME_SHAPE_LAYOUT.Name)
                {
                    CT_ShapeLayout sl=CT_ShapeLayout.Parse(xmltext);
                    _items.Add(sl);
                }
                else if (nd.LocalName == QNAME_SHAPE_TYPE.Name)
                {
                    CT_Shapetype st = CT_Shapetype.Parse(xmltext);
                    _shapeTypeId = st.id;
                    _items.Add(st);
                }
                else if (nd.LocalName == QNAME_SHAPE.Name)
                {
                    CT_Shape shape = CT_Shape.Parse(xmltext);
                    String id = shape.id;
                    if (id != null)
                    {
                        MatchCollection m = ptrn_shapeId.Matches(id);
                        if (m.Count>0) 
                            _shapeId = Math.Max(_shapeId, int.Parse(m[0].Groups[1].Value));
                    }
                    _items.Add(shape);
                }
                else
                {
                    _items.Add(nd);
                }

            }
        }

        internal ArrayList GetItems()
        {
            return _items;
        }

        internal void Write(Stream out1)
        {
            XmlWriter xw = XmlWriter.Create(out1);
            xw.WriteStartElement("xml");
            xw.WriteAttributeString("xmlns", "v",null, "urn:schemas-microsoft-com:vml");
            xw.WriteAttributeString("xmlns", "o",null, "urn:schemas-microsoft-com:office:office");
            xw.WriteAttributeString("xmlns", "x",null, "urn:schemas-microsoft-com:office:excel");
           
            for (int i = 0; i < _items.Count; i++)
            {
                object xc = _items[i];
                if (xc is XmlNode)
                {
                    xw.WriteRaw(((XmlNode)xc).OuterXml.Replace(" xmlns:v=\"urn:schemas-microsoft-com:vml\"", "").Replace(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"", "").Replace(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"", ""));
                }
                else
                {
                    xw.WriteRaw(xc.ToString().Replace(" xmlns:v=\"urn:schemas-microsoft-com:vml\"", "").Replace(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"", "").Replace(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"", ""));
                }
            }

            xw.WriteEndElement();
            xw.Flush();
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

            CT_Shapetype shapetype = new CT_Shapetype();
            _shapeTypeId = "_xssf_cell_comment";
            shapetype.id= _shapeTypeId;
            shapetype.coordsize="21600,21600";
            shapetype.spt=202;
            shapetype.path2 = ("m,l,21600r21600,l21600,xe");
            shapetype.AddNewStroke().joinstyle = (ST_StrokeJoinStyle.miter);
            CT_Path path = shapetype.AddNewPath();
            path.gradientshapeok = ST_TrueFalse.t;
            path.connecttype=(ST_ConnectType.rect);
            _items.Add(shapetype);
        }

        internal CT_Shape newCommentShape()
        {
            CT_Shape shape = new CT_Shape();

            shape.id = ("_x0000_s" + (++_shapeId));
            shape.type =("#" + _shapeTypeId);
            shape.style=("position:absolute; visibility:hidden");
            shape.fillcolor = ("#ffffe1");
            shape.insetmode = (ST_InsetMode.auto);
            shape.AddNewFill().color=("#ffffe1");
            CT_Shadow shadow = shape.AddNewShadow();
            shadow.on= ST_TrueFalse.t;
            shadow.color = "black";
            shadow.obscured = ST_TrueFalse.t;
            shape.AddNewPath().connecttype = (ST_ConnectType.none);
            shape.AddNewTextbox().style = ("mso-direction-alt:auto");
            CT_ClientData cldata = shape.AddNewClientData();
            cldata.ObjectType=ST_ObjectType.Note;
            cldata.AddNewMoveWithCells();
            cldata.AddNewSizeWithCells();
            cldata.AddNewAnchor("1, 15, 0, 2, 3, 15, 3, 16");
            cldata.AddNewAutoFill(ST_TrueFalseBlank.@false);
            cldata.AddNewRow(0);
            cldata.AddNewColumn(0);
            _items.Add(shape);

            return shape;
        }

        /**
         * Find a shape with ClientData of type "NOTE" and the specified row and column
         *
         * @return the comment shape or <code>null</code>
         */
        internal CT_Shape FindCommentShape(int row, int col)
        {
            foreach (object itm in _items)
            {
                if (itm is CT_Shape)
                {
                    CT_Shape sh = (CT_Shape)itm;
                    if (sh.sizeOfClientDataArray() > 0)
                    {
                        CT_ClientData cldata = sh.GetClientDataArray(0);
                        if (cldata.ObjectType == ST_ObjectType.Note)
                        {
                            int crow = cldata.GetRowArray(0);
                            int ccol = cldata.GetColumnArray(0);
                            if (crow == row && ccol == col)
                            {
                                return sh;
                            }
                        }
                    }
                }
            }
            return null;
        }

        internal bool RemoveCommentShape(int row, int col)
        {
            CT_Shape shape = FindCommentShape(row, col);
            if(shape == null)
                return false;
             _items.Remove(shape);
             return true;
        }
    }
}

