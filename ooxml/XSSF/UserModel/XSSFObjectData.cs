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

namespace NPOI.XSSF.UserModel

{


    using NPOI;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.OpenXmlFormats.Dml.Spreadsheet;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using System;
    using System.IO;
    using System.Runtime.ConstrainedExecution;
    using System.Xml.Linq;


    /// <summary>
    /// Represents binary object (i.e. OLE) data stored in the file.  Eg. A GIF, JPEG etc...
    /// </summary>
    public class XSSFObjectData : XSSFSimpleShape, IObjectData
    {
        private static POILogger LOG = POILogFactory.GetLogger(typeof(XSSFObjectData));

        /// <summary>
        /// A default instance of CTShape used for creating new shapes.
        /// </summary>
        private static CT_Shape prototype = null;

        private CT_OleObject oleObject;

        public XSSFObjectData(XSSFDrawing drawing, CT_Shape ctShape)
            : base(drawing, ctShape)
        {
            ;
        }

        /// <summary>
        /// Prototype with the default structure of a new auto-shape.
        /// </summary>
        /// <summary>
        /// Prototype with the default structure of a new auto-shape.
        /// </summary>
        public new static CT_Shape Prototype()
        {
            String drawNS = "http://schemas.microsoft.com/office/drawing/2010/main";

            if (prototype == null)
            {
                CT_Shape shape = new CT_Shape();

                CT_ShapeNonVisual nv = shape.AddNewNvSpPr();
                OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualDrawingProps nvp = nv.AddNewCNvPr();
                nvp.id = (/*setter*/1);
                nvp.name = (/*setter*/"Shape 1");
                //            nvp.Hidden=(/*setter*/true);
                CT_OfficeArtExtensionList extLst = nvp.AddNewExtLst();
                // https://msdn.microsoft.com/en-us/library/dd911027(v=office.12).aspx
                CT_OfficeArtExtension ext = extLst.AddNewExt();
                ext.uri = (/*setter*/"{63B3BB69-23CF-44E3-9099-C40C66FF867C}");
                //XmlCursor cur = ext.NewCursor();
                //cur.ToEndToken();
                //cur.BeginElement(new QName(drawNS, "compatExt", "a14"));
                //cur.InsertNamespace("a14", drawNS);
                //cur.InsertAttributeWithValue("spid", "_x0000_s1");
                //cur.Dispose();
                ext.Any = "<a14:compatExt xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" spid=\"_x0000_s1\"/>";

                nv.AddNewCNvSpPr();

                OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties sp = shape.AddNewSpPr();
                CT_Transform2D t2d = sp.AddNewXfrm();
                CT_PositiveSize2D p1 = t2d.AddNewExt();
                p1.cx = (/*setter*/0);
                p1.cy = (/*setter*/0);
                CT_Point2D p2 = t2d.AddNewOff();
                p2.x = (/*setter*/0);
                p2.y = (/*setter*/0);

                CT_PresetGeometry2D geom = sp.AddNewPrstGeom();
                geom.prst = (/*setter*/ST_ShapeType.rect);
                geom.AddNewAvLst();

                prototype = shape;
            }
            return prototype;
        }
        public String OLE2ClassName
        {
            get
            {
                return GetOleObject().progId;
            }
            
        }

        /// <summary>
        /// </summary>
        /// <return>CTOleObject associated with the shape/// </return>
        public CT_OleObject GetOleObject()
        {
            if (oleObject == null)
            {
                long shapeId = GetCTShape().nvSpPr.cNvPr.id;
                oleObject = GetSheet().ReadOleObject(shapeId);
                if (oleObject == null)
                {
                    throw new POIXMLException("Ole object not found in sheet Container - it's probably a control element");
                }
            }
            return oleObject;
        }
        public byte[] ObjectData
        {
            get
            {
                Stream is1 = GetObjectPart().GetInputStream();
                MemoryStream bos = new MemoryStream();
                IOUtils.Copy(is1, bos);
                is1.Close();
                return bos.ToArray();
            }
        }

        /// <summary>
        /// </summary>
        /// <return>package part of the object data/// </return>
        public PackagePart GetObjectPart()
        {
            if (!GetOleObject().IsSetId())
            {
                throw new POIXMLException("Invalid ole object found in sheet Container");
            }
            POIXMLDocumentPart pdp = GetSheet().GetRelationById(GetOleObject().id);
            return (pdp == null) ? null : pdp.GetPackagePart();
        }
        public bool HasDirectoryEntry()
        {
            Stream is1 = null;
            try
            {
                is1 = GetObjectPart().GetInputStream();// as InputStream;

                // If Clearly doesn't do mark/reset, wrap up
                //if (!is1.MarkSupported())
                //{
                //    is1 = new PushbackInputStream(is1, 8);
                //}

                // Ensure that there is at least some data there
                byte[] header8 = IOUtils.PeekFirstNBytes(is1, 8);

                // Try to create
                return NPOIFSFileSystem.HasPOIFSHeader(header8);
            }
            catch (IOException e)
            {
                LOG.Log(POILogger.WARN, "can't determine if directory entry exists", e);
                return false;
            }
            finally
            {
                IOUtils.CloseQuietly(is1);
            }
        }

        public DirectoryEntry Directory
        {
            get
            {
                Stream is1 = null;
                try
                {
                    is1 = GetObjectPart().GetInputStream();
                    return new POIFSFileSystem(is1).Root;
                }
                finally
                {
                    IOUtils.CloseQuietly(is1);
                }
            }
        }

        /// <summary>
        /// The filename of the embedded image
        /// </summary>
        public String FileName
        {
            get
            {
                return GetObjectPart().PartName.Name;
            }
        }

        protected XSSFSheet GetSheet()
        {
            return (XSSFSheet)GetDrawing().GetParent();
        }
        public IPictureData PictureData
        {
            get
            {
                var oleObj = GetOleObject();
                if(oleObj.objectPr!=null && !string.IsNullOrEmpty(oleObj.objectPr.id))
                {
                    return (XSSFPictureData) GetSheet().GetRelationById(oleObj.objectPr.id);
                }
                else
                    return null;

                //XmlCursor cur = GetOleObject().newCursor();
                //try
                //{
                //    if (cur.ToChild(XSSFRelation.NS_SPREADSHEETML, "objectPr"))
                //    {
                //        String blipId = cur.GetAttributeText(new QName(PackageRelationshipTypes.CORE_PROPERTIES_ECMA376_NS, "id"));
                //        return (XSSFPictureData)getSheet().GetRelationById(blipId);
                //    }
                //    return null;
                //}
                //finally
                //{
                //    cur.Dispose();
                //}
            }
        }
    }
}

