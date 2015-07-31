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

using System;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml.Spreadsheet; // http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using System.Collections.Generic;
using System.Xml;
using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Util;
using NPOI.Util;


namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a SpreadsheetML Drawing
     *
     * @author Yegor Kozlov
     */
    public class XSSFDrawing : POIXMLDocumentPart, IDrawing
    {
        public const String NAMESPACE_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public const String NAMESPACE_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        /**
         * Root element of the SpreadsheetML Drawing part
         */
        private NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing drawing = NewDrawing();
       // private bool isNew = true; not used so far
        private long numOfGraphicFrames = 0L;

        /**
         * Create a new SpreadsheetML Drawing
         *
         * @see NPOI.xssf.usermodel.XSSFSheet#CreateDrawingPatriarch()
         */
        public XSSFDrawing()
            : base()
        {
            drawing = NewDrawing();
        }
        /**
         * Construct a SpreadsheetML Drawing from a namespace part
         *
         * @param part the namespace part holding the Drawing data,
         * the content type must be <code>application/vnd.openxmlformats-officedocument.Drawing+xml</code>
         * @param rel  the namespace relationship holding this Drawing,
         * the relationship type must be http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing
         */
        internal XSSFDrawing(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            XmlDocument xmldoc = ConvertStreamToXml(part.GetInputStream());
            drawing = NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing.Parse(xmldoc, NamespaceManager);
        }

        /**
         * Construct a new CT_Drawing bean. By default, it's just an empty placeholder for Drawing objects
         *
         * @return a new CT_Drawing bean
         */
        private static NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_Drawing NewDrawing()
        {
            return new CT_Drawing();
        }

        /**
         * Return the underlying CT_Drawing bean, the root element of the SpreadsheetML Drawing part.
         *
         * @return the underlying CT_Drawing bean
         */
        public CT_Drawing GetCTDrawing()
        {
            return drawing;
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            drawing.Save(out1);
            out1.Close();
        }

        public IClientAnchor CreateAnchor(int dx1, int dy1, int dx2, int dy2,
                int col1, int row1, int col2, int row2)
        {
            return new XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
        }

        /**
         * Constructs a textbox under the Drawing.
         *
         * @param anchor    the client anchor describes how this group is attached
         *                  to the sheet.
         * @return      the newly Created textbox.
         */
        public XSSFTextBox CreateTextbox(IClientAnchor anchor)
        {
            long shapeId = newShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Shape ctShape = ctAnchor.AddNewSp();
            ctShape.Set(XSSFSimpleShape.GetPrototype());
            ctShape.nvSpPr.cNvPr.id=(uint)shapeId;
            XSSFTextBox shape = new XSSFTextBox(this, ctShape);
            shape.anchor = (XSSFClientAnchor)anchor;
            return shape;

        }

        /**
         * Creates a picture.
         *
         * @param anchor    the client anchor describes how this picture is attached to the sheet.
         * @param pictureIndex the index of the picture in the workbook collection of pictures,
         *   {@link NPOI.xssf.usermodel.XSSFWorkbook#getAllPictures()} .
         *
         * @return  the newly Created picture shape.
         */
        public IPicture CreatePicture(XSSFClientAnchor anchor, int pictureIndex)
        {
            PackageRelationship rel = AddPictureReference(pictureIndex);

            long shapeId = newShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Picture ctShape = ctAnchor.AddNewPic();
            ctShape.Set(XSSFPicture.Prototype());

            ctShape.nvPicPr.cNvPr.id = (uint)shapeId;
            ctShape.nvPicPr.cNvPr.name = "Picture " + shapeId;

            XSSFPicture shape = new XSSFPicture(this, ctShape);
            shape.anchor = anchor;
            shape.SetPictureReference(rel);
            return shape;
        }

        public IPicture CreatePicture(IClientAnchor anchor, int pictureIndex)
        {
            return CreatePicture((XSSFClientAnchor)anchor, pictureIndex);
        }
        /// <summary>
        /// Creates a chart.
        /// </summary>
        /// <param name="anchor">the client anchor describes how this chart is attached to</param>
        /// <returns>the newly created chart</returns>
        public IChart CreateChart(IClientAnchor anchor)
        {
            int chartNumber = GetPackagePart().Package.
                GetPartsByContentType(XSSFRelation.CHART.ContentType).Count + 1;

            XSSFChart chart = (XSSFChart)CreateRelationship(
                    XSSFRelation.CHART, XSSFFactory.GetInstance(), chartNumber);
            String chartRelId = chart.GetPackageRelationship().Id;

            XSSFGraphicFrame frame = CreateGraphicFrame((XSSFClientAnchor)anchor);
            frame.SetChart(chart, chartRelId);

            return chart;
        }

        //public XSSFChart CreateChart(IClientAnchor anchor)
        //{
        //    return CreateChart((XSSFClientAnchor)anchor);
        //}

        /**
         * Add the indexed picture to this Drawing relations
         *
         * @param pictureIndex the index of the picture in the workbook collection of pictures,
         *   {@link NPOI.xssf.usermodel.XSSFWorkbook#getAllPictures()} .
         */
        internal PackageRelationship AddPictureReference(int pictureIndex)
        {
            XSSFWorkbook wb = (XSSFWorkbook)GetParent().GetParent();
            XSSFPictureData data = (XSSFPictureData)wb.GetAllPictures()[pictureIndex];
            PackagePartName ppName = data.GetPackagePart().PartName;
            PackageRelationship rel = GetPackagePart().AddRelationship(ppName, TargetMode.Internal, XSSFRelation.IMAGES.Relation);
            AddRelation(rel.Id, new XSSFPictureData(data.GetPackagePart(), rel));
            return rel;
        }

        /**
         * Creates a simple shape.  This includes such shapes as lines, rectangles,
         * and ovals.
         *
         * @param anchor    the client anchor describes how this group is attached
         *                  to the sheet.
         * @return  the newly Created shape.
         */
        public XSSFSimpleShape CreateSimpleShape(XSSFClientAnchor anchor)
        {
            long shapeId = newShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Shape ctShape = ctAnchor.AddNewSp();
            ctShape.Set(XSSFSimpleShape.GetPrototype());
            ctShape.nvSpPr.cNvPr.id=(uint)(shapeId);
            XSSFSimpleShape shape = new XSSFSimpleShape(this, ctShape);
            shape.anchor = anchor;
            return shape;
        }

        /**
         * Creates a simple shape.  This includes such shapes as lines, rectangles,
         * and ovals.
         *
         * @param anchor    the client anchor describes how this group is attached
         *                  to the sheet.
         * @return  the newly Created shape.
         */
        public XSSFConnector CreateConnector(XSSFClientAnchor anchor)
        {
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Connector ctShape = ctAnchor.AddNewCxnSp();
            ctShape.Set(XSSFConnector.Prototype());

            XSSFConnector shape = new XSSFConnector(this, ctShape);
            shape.anchor = anchor;
            return shape;
        }

        /**
         * Creates a simple shape.  This includes such shapes as lines, rectangles,
         * and ovals.
         *
         * @param anchor    the client anchor describes how this group is attached
         *                  to the sheet.
         * @return  the newly Created shape.
         */
        public XSSFShapeGroup CreateGroup(XSSFClientAnchor anchor)
        {
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_GroupShape ctGroup = ctAnchor.AddNewGrpSp();
            ctGroup.Set(XSSFShapeGroup.Prototype());

            XSSFShapeGroup shape = new XSSFShapeGroup(this, ctGroup);
            shape.anchor = anchor;
            return shape;
        }

        /**
         * Creates a comment.
         * @param anchor the client anchor describes how this comment is attached
         *               to the sheet.
         * @return the newly Created comment.
         */
        public IComment CreateCellComment(IClientAnchor anchor)
        {
            XSSFClientAnchor ca = (XSSFClientAnchor)anchor;
            XSSFSheet sheet = (XSSFSheet)GetParent();

            //create comments and vmlDrawing parts if they don't exist
            CommentsTable comments = sheet.GetCommentsTable(true);
            XSSFVMLDrawing vml = sheet.GetVMLDrawing(true);
            NPOI.OpenXmlFormats.Vml.CT_Shape vmlShape = vml.newCommentShape();
            if (ca.IsSet())
            {
                // convert offsets from emus to pixels since we get a DrawingML-anchor
                // but create a VML Drawing
                int dx1Pixels = ca.Dx1 / Units.EMU_PER_PIXEL;
                int dy1Pixels = ca.Dy1 / Units.EMU_PER_PIXEL;
                int dx2Pixels = ca.Dx2 / Units.EMU_PER_PIXEL;
                int dy2Pixels = ca.Dy2 / Units.EMU_PER_PIXEL;
                String position =
                        ca.Col1 + ", " + dx1Pixels + ", " +
                        ca.Row1 + ", " + dy1Pixels + ", " +
                        ca.Col2 + ", " + dx2Pixels + ", " +
                        ca.Row2 + ", " + dy2Pixels;
                vmlShape.GetClientDataArray(0).SetAnchorArray(0, position);
            }
            String ref1 = new CellReference(ca.Row1, ca.Col1).FormatAsString();
            if (comments.FindCellComment(ref1) != null)
            {
                throw new ArgumentException("Multiple cell comments in one cell are not allowed, cell: " + ref1);
            }

            return new XSSFComment(comments, comments.NewComment(ref1), vmlShape);
        }

        /**
         * Creates a new graphic frame.
         *
         * @param anchor    the client anchor describes how this frame is attached
         *                  to the sheet
         * @return  the newly Created graphic frame
         */
        private XSSFGraphicFrame CreateGraphicFrame(XSSFClientAnchor anchor)
        {
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_GraphicalObjectFrame ctGraphicFrame = ctAnchor.AddNewGraphicFrame();
            ctGraphicFrame.Set(XSSFGraphicFrame.Prototype());

            long frameId = numOfGraphicFrames++;
            XSSFGraphicFrame graphicFrame = new XSSFGraphicFrame(this, ctGraphicFrame);
            graphicFrame.Anchor = anchor;
            graphicFrame.Id = frameId;
            graphicFrame.Name = "Diagramm" + frameId;
            return graphicFrame;
        }

        /**
         * Returns all charts in this Drawing.
         */
        public List<XSSFChart> GetCharts()
        {
            List<XSSFChart> charts = new List<XSSFChart>();
            foreach (POIXMLDocumentPart part in GetRelations())
            {
                if (part is XSSFChart)
                {
                    charts.Add((XSSFChart)part);
                }
            }
            return charts;
        }

        /**
         * Create and Initialize a CT_TwoCellAnchor that anchors a shape against top-left and bottom-right cells.
         *
         * @return a new CT_TwoCellAnchor
         */
        private CT_TwoCellAnchor CreateTwoCellAnchor(IClientAnchor anchor)
        {
            CT_TwoCellAnchor ctAnchor = drawing.AddNewTwoCellAnchor();
            XSSFClientAnchor xssfanchor = (XSSFClientAnchor)anchor;
            ctAnchor.from = (xssfanchor.From);
            ctAnchor.to = (xssfanchor.To);
            ctAnchor.AddNewClientData();
            xssfanchor.To = ctAnchor.to;
            xssfanchor.From = ctAnchor.from;
            ST_EditAs aditAs;
            switch (anchor.AnchorType)
            {
                case AnchorType.DontMoveAndResize: 
                    aditAs = ST_EditAs.absolute; break;
                case AnchorType.MoveAndResize: 
                    aditAs = ST_EditAs.twoCell; break;
                case AnchorType.MoveDontResize: 
                    aditAs = ST_EditAs.oneCell; break;
                default: 
                    aditAs = ST_EditAs.oneCell;
                    break;
            }
            ctAnchor.editAs = aditAs;
            ctAnchor.editAsSpecified = true;
            return ctAnchor;
        }

        private long newShapeId()
        {
            return drawing.SizeOfTwoCellAnchorArray() + 1;
        }

        #region IDrawing Members

        public bool ContainsChart()
        {
            throw new NotImplementedException();
        }

        #endregion

        /**
     *
     * @return list of shapes in this drawing
     */
        public List<XSSFShape> GetShapes()
        {
            List<XSSFShape> lst = new List<XSSFShape>();
            foreach (IEG_Anchor anchor in drawing.CellAnchors)
            {
                XSSFShape shape = null;
                if (anchor.picture != null)
                {
                    shape = new XSSFPicture(this, anchor.picture);
                }
                else if (anchor.connector != null)
                {
                    shape = new XSSFConnector(this, anchor.connector);
                }
                else if (anchor.groupShape != null)
                {
                    shape = new XSSFShapeGroup(this, anchor.groupShape);
                }
                else if (anchor.graphicFrame != null)
                {
                    shape = new XSSFGraphicFrame(this, anchor.graphicFrame);
                }
                else if (anchor.sp != null)
                {
                   shape = new XSSFSimpleShape(this, anchor.sp);
                }
                if (shape != null)
                {
                    shape.anchor = GetAnchorFromIEGAnchor(anchor);
                    lst.Add(shape);
                }
            }
            //foreach (XmlNode obj in xmldoc.SelectNodes("./*/*/*"))
            //{
            //    XSSFShape shape = null;
            //    if (obj.LocalName == "sp")
            //    {
            //        shape = new XSSFSimpleShape(this, obj);
            //    }
            //    else if (obj.LocalName == "pic")
            //    {
            //        shape = new XSSFPicture(this, obj);
            //    }
            //    else if (obj.LocalName == "cxnSp")
            //    {
            //        shape = new XSSFConnector(this, obj);
            //    }
            //    //    else if (obj is CT_GraphicalObjectFrame) shape = new XSSFGraphicFrame(this, (CT_GraphicalObjectFrame)obj);
            //    //    else if (obj is CT_GroupShape) shape = new XSSFShapeGroup(this, (CT_GroupShape)obj);
            //    if (shape != null)
            //    {
            //        shape.anchor = GetAnchorFromParent(obj);
            //        lst.Add(shape);
            //    }
            //}
            return lst;
        }
        private XSSFAnchor GetAnchorFromIEGAnchor(IEG_Anchor ctAnchor)
        {
            CT_Marker ctFrom=null, ctTo=null;
            if (ctAnchor is CT_TwoCellAnchor)
            {
                ctFrom = ((CT_TwoCellAnchor)ctAnchor).from;
                ctTo = ((CT_TwoCellAnchor)ctAnchor).to;
            }
            else if (ctAnchor is CT_OneCellAnchor)
            {
                ctFrom = ((CT_OneCellAnchor)ctAnchor).from;
            }
            XSSFAnchor anchor = new XSSFClientAnchor(ctFrom, ctTo);
            return anchor;
        }
        private XSSFAnchor GetAnchorFromParent(XmlNode obj)
        {
            XSSFAnchor anchor = null;
            XmlNode parentNode = obj.ParentNode;
            XmlNode fromNode = parentNode.SelectSingleNode("xdr:from", POIXMLDocumentPart.NamespaceManager);
            if(fromNode==null)
                throw new InvalidDataException("xdr:from node is missing");
            CT_Marker ctFrom = CT_Marker.Parse(fromNode, POIXMLDocumentPart.NamespaceManager);
            XmlNode toNode = parentNode.SelectSingleNode("xdr:to", POIXMLDocumentPart.NamespaceManager);
            CT_Marker ctTo=null;
            if (toNode != null)
            {
                ctTo = CT_Marker.Parse(toNode, POIXMLDocumentPart.NamespaceManager);
            }
            anchor = new XSSFClientAnchor(ctFrom, ctTo);
            return anchor;
        }
    }
}

