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
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections;
using NPOI.OpenXml4Net.Exceptions;

namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a SpreadsheetML Drawing
     *
     * @author Yegor Kozlov
     */
    public class XSSFDrawing : POIXMLDocumentPart, IDrawing<IShape>
    {
        public static String NAMESPACE_A = XSSFRelation.NS_DRAWINGML;
        public static String NAMESPACE_C = XSSFRelation.NS_CHART;

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
        internal XSSFDrawing(PackagePart part)
            : base(part)
        {
            XmlDocument xmldoc = ConvertStreamToXml(part.GetInputStream());
            drawing = CT_Drawing.Parse(xmldoc, NamespaceManager);
        }

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public XSSFDrawing(PackagePart part, PackageRelationship rel)
            : this(part)
        {
            
        }
        /**
         * Construct a new CT_Drawing bean. By default, it's just an empty placeholder for Drawing objects
         *
         * @return a new CT_Drawing bean
         */
        private static CT_Drawing NewDrawing()
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
            long shapeId = NewShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Shape ctShape = ctAnchor.AddNewSp();
            ctShape.Set(XSSFSimpleShape.Prototype());
            ctShape.nvSpPr.cNvPr.id=(uint)shapeId;
            ctShape.spPr.xfrm.off.x = ((XSSFClientAnchor)(anchor)).left;
            ctShape.spPr.xfrm.off.y = ((XSSFClientAnchor)(anchor)).top;
            ctShape.spPr.xfrm.ext.cx = ((XSSFClientAnchor)(anchor)).width;
            ctShape.spPr.xfrm.ext.cy = ((XSSFClientAnchor)(anchor)).height;
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

            long shapeId = NewShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Picture ctShape = ctAnchor.AddNewPic();
            ctShape.Set(XSSFPicture.Prototype());

            ctShape.nvPicPr.cNvPr.id = (uint)shapeId;
            ctShape.spPr.xfrm.off.x = anchor.left;
            ctShape.spPr.xfrm.off.y = anchor.top;
            ctShape.spPr.xfrm.ext.cx = anchor.width;
            ctShape.spPr.xfrm.ext.cy = anchor.height;
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
            int chartNumber = GetNewChartNumber();

            RelationPart rp = CreateRelationship(
            XSSFRelation.CHART, XSSFFactory.GetInstance(), chartNumber, false);
            XSSFChart chart = rp.DocumentPart as XSSFChart;
            String chartRelId = rp.Relationship.Id;

            XSSFGraphicFrame frame = CreateGraphicFrame((XSSFClientAnchor)anchor);
            frame.SetChart(chart, chartRelId);

            return chart;
        }

        /// <summary>
        /// Removes chart.
        /// </summary>
        /// <param name="chart">The chart to be removed.</param>
        public void RemoveChart(XSSFChart chart)
        {
            CT_Drawing ctDrawing = GetCTDrawing();
            XSSFGraphicFrame frame = chart.GetGraphicFrame();
            CT_GraphicalObjectFrame internalFrame = frame.GetCTGraphicalObjectFrame();
            int anchorIndex = ctDrawing.CellAnchors.FindIndex(anchor => anchor.graphicFrame == internalFrame);

            if (anchorIndex != -1)
            {
                ctDrawing.CellAnchors.RemoveAt(anchorIndex);

                foreach (var part in GetRelations().Where(part => part is XSSFChart && part == chart))
                {
                    RemoveRelation(part);
                }
            }
        }

        private int GetNewChartNumber()
        {
            List<PackagePart> existingCharts = GetPackagePart().Package.
                GetPartsByContentType(XSSFRelation.CHART.ContentType);
            HashSet<int> numbers = new HashSet<int>();

            foreach (PackagePart chart in existingCharts)
            {
                var match = Regex.Match(chart.PartName.Name, @"\d+");

                if (match.Success)
                {
                    numbers.Add(int.Parse(match.Value));
                }
            }

            var newChartNumber = 1;

            while (numbers.Contains(newChartNumber))
            {
                newChartNumber++;
            }

            return newChartNumber;
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
            XSSFPictureData pic = new XSSFPictureData(data.GetPackagePart());
            RelationPart rp = AddRelation(null, XSSFRelation.IMAGES, pic);
            return rp.Relationship;
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
            long shapeId = NewShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Shape ctShape = ctAnchor.AddNewSp();
            ctShape.Set(XSSFSimpleShape.Prototype());
            ctShape.nvSpPr.cNvPr.id=(uint)(shapeId);
            ctShape.spPr.xfrm.off.x = anchor.left;
            ctShape.spPr.xfrm.off.y = anchor.top;
            ctShape.spPr.xfrm.ext.cx = anchor.width;
            ctShape.spPr.xfrm.ext.cy = anchor.height;
            XSSFSimpleShape shape = new XSSFSimpleShape(this, ctShape);
            shape.anchor = anchor;
            shape.cellanchor = ctAnchor;

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
            long shapeId = NewShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Connector ctShape = ctAnchor.AddNewCxnSp();
            ctShape.Set(XSSFConnector.Prototype());
            ctShape.nvCxnSpPr.cNvPr.id = (uint)(shapeId);
            ctShape.spPr.xfrm.off.x = anchor.left;
            ctShape.spPr.xfrm.off.y = anchor.top;
            ctShape.spPr.xfrm.ext.cx = anchor.width;
            ctShape.spPr.xfrm.ext.cy = anchor.height;

            XSSFConnector shape = new XSSFConnector(this, ctShape);
            shape.anchor = anchor;
            shape.cellanchor = ctAnchor;

            return shape;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="anchor"></param>
        /// <param name="coords"></param>
        /// <returns></returns>
        public XSSFFreeform CreateFreeform(
              SS.UserModel.ISheet Sheet
            , BuildFreeForm BFF
        ) {
            var anchor = new XSSFClientAnchor(Sheet, (int)BFF.Left, (int)BFF.Top
                                                   , (int)BFF.Rigth, (int)BFF.Bottom);
            long shapeId = NewShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_Shape ctShape = ctAnchor.AddNewSp();
            ctShape.Set(XSSFFreeform.Prototype());
            ctShape.nvSpPr.cNvPr.id = (uint) (shapeId);
            var freeform = new XSSFFreeform(this, ctShape);
            freeform.anchor = anchor;
            freeform.cellanchor = ctAnchor;
                
            freeform.Build(BFF);

            return freeform;
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
            long shapeId = NewShapeId();
            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor(anchor);
            CT_GroupShape ctGroup = ctAnchor.AddNewGrpSp();
            ctGroup.Set(XSSFShapeGroup.Prototype());
            ctGroup.nvGrpSpPr.cNvPr.id = (uint)(shapeId);
            ctGroup.grpSpPr.xfrm.off.x = anchor.left;
            ctGroup.grpSpPr.xfrm.off.y = anchor.top;
            ctGroup.grpSpPr.xfrm.ext.cx = anchor.width;
            ctGroup.grpSpPr.xfrm.ext.cy = anchor.height;
            ctGroup.grpSpPr.xfrm.chOff.x = ctGroup.grpSpPr.xfrm.off.x;
            ctGroup.grpSpPr.xfrm.chOff.y = ctGroup.grpSpPr.xfrm.off.y;
            ctGroup.grpSpPr.xfrm.chExt.cx =ctGroup.grpSpPr.xfrm.ext.cx;
            ctGroup.grpSpPr.xfrm.chExt.cy =ctGroup.grpSpPr.xfrm.ext.cy;

            XSSFShapeGroup shape = new XSSFShapeGroup(this, ctGroup);
            shape.anchor = anchor;
            shape.cellanchor = ctAnchor;

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
            CellAddress ref1 = new CellAddress(ca.Row1, ca.Col1);
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
            ctGraphicFrame.xfrm = CreateXfrm(anchor);

            long frameId = numOfGraphicFrames++;
            XSSFGraphicFrame graphicFrame = new XSSFGraphicFrame(this, ctGraphicFrame);
            graphicFrame.Anchor = anchor;
            graphicFrame.Id = frameId;
            graphicFrame.Name = "Diagramm" + frameId;
            graphicFrame.cellanchor = ctAnchor;

            return graphicFrame;
        }

        public IObjectData CreateObjectData(IClientAnchor anchor, int storageId, int pictureIndex)
        {
            XSSFSheet sh = Sheet;
            PackagePart sheetPart = sh.GetPackagePart();

            /*
             * The shape id of the ole object seems to be a legacy shape id.
             * 
             * see 5.3.2.1 legacyDrawing (Legacy Drawing Object):
             * Legacy Shape ID that is unique throughout the entire document.
             * Legacy shape IDs should be assigned based on which portion of the document the
             * drawing resides on. The assignment of these ids is broken down into clusters of
             * 1024 values. The first cluster is 1-1024, the second 1025-2048 and so on.
             *
             * Ole shapes seem to start with 1025 on the first sheet ...
             * and not sure, if the ids need to be reindexed when sheets are removed
             * or more than 1024 shapes are on a given sheet (see #51332 for a similar issue)
             */
            XSSFSheet sheet = Sheet;
            XSSFWorkbook wb = sheet.Workbook as XSSFWorkbook;
            int sheetIndex = wb.GetSheetIndex(sheet);
            long shapeId = (sheetIndex+1)*1024 + NewShapeId();

            // add reference to OLE part
            PackagePartName olePN;
            try
            {
                olePN = PackagingUriHelper.CreatePartName("/xl/embeddings/oleObject"+storageId+".bin");
            }
            catch(InvalidFormatException e)
            {
                throw new POIXMLException(e);
            }
            PackageRelationship olePR = sheetPart.AddRelationship( olePN, TargetMode.Internal, POIXMLDocument.OLE_OBJECT_REL_TYPE );

            // add reference to image part
            XSSFPictureData imgPD = sh.Workbook.GetAllPictures()[pictureIndex] as XSSFPictureData;
            PackagePartName imgPN = imgPD.GetPackagePart().PartName;
            PackageRelationship imgSheetPR = sheetPart.AddRelationship( imgPN, TargetMode.Internal, PackageRelationshipTypes.IMAGE_PART );
            PackageRelationship imgDrawPR = GetPackagePart().AddRelationship( imgPN, TargetMode.Internal, PackageRelationshipTypes.IMAGE_PART );


            // add OLE part metadata to sheet
            NPOI.OpenXmlFormats.Spreadsheet.CT_Worksheet cwb = sh.GetCTWorksheet();
            NPOI.OpenXmlFormats.Spreadsheet.CT_OleObjects oo = cwb.IsSetOleObjects() ? cwb.oleObjects : cwb.AddNewOleObjects();

            NPOI.OpenXmlFormats.Spreadsheet.CT_OleObject ole1 = oo.AddNewOleObject();
            ole1.progId = "Package";
            ole1.shapeId = (uint) shapeId;
            ole1.id = olePR.Id;

            //XmlCursor cur1 = ole1.newCursor();
            //cur1.toEndToken();
            //cur1.beginElement("objectPr", XSSFRelation.NS_SPREADSHEETML);
            //cur1.insertAttributeWithValue("id", PackageRelationshipTypes.CORE_PROPERTIES_ECMA376_NS, imgSheetPR.Id);
            //cur1.insertAttributeWithValue("defaultSize", "0");
            //cur1.beginElement("anchor", XSSFRelation.NS_SPREADSHEETML);
            //cur1.insertAttributeWithValue("moveWithCells", "1");

            ole1.objectPr.id = imgSheetPR.Id;
            ole1.objectPr.defaultSize = false;
            ole1.objectPr.anchor.moveWithCells = true;

            CT_TwoCellAnchor ctAnchor = CreateTwoCellAnchor((XSSFClientAnchor)anchor);

            //XmlCursor cur2 = ctAnchor.newCursor();
            //cur2.copyXmlContents(cur1);
            //cur2.dispose();

            //cur1.toParent();
            //cur1.toFirstChild();
            //cur1.setName(new QName(XSSFRelation.NS_SPREADSHEETML, "from"));
            //cur1.toNextSibling();
            //cur1.setName(new QName(XSSFRelation.NS_SPREADSHEETML, "to"));

            //cur1.dispose();

            // add a new shape and link OLE & image part
            CT_Shape ctShape = ctAnchor.AddNewSp();
            ctShape.Set(XSSFObjectData.Prototype());
            ctShape.spPr.xfrm = (CreateXfrm((XSSFClientAnchor) anchor));

            // workaround for not having the vmlDrawing filled
            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_BlipFillProperties blipFill = ctShape.spPr.AddNewBlipFill();
            blipFill.AddNewBlip().embed = imgDrawPR.Id;
            blipFill.AddNewStretch().AddNewFillRect();

            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualDrawingProps cNvPr = ctShape.nvSpPr.cNvPr;
            cNvPr.id = (uint)shapeId;
            cNvPr.name = "Object "+shapeId;

            //XmlCursor extCur = cNvPr.getExtLst().getExtArray(0).newCursor();
            //extCur.toFirstChild();
            //extCur.setAttributeText(new QName("spid"), "_x0000_s"+shapeId);
            //extCur.dispose();
            var ext = cNvPr.extLst.ext[0];
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(ext.Any);
            doc.FirstChild.Attributes["spid"].Value = "_x0000_s"+shapeId;

            ext.Any = doc.InnerXml;

            XSSFObjectData shape = new XSSFObjectData(this, ctShape);
            shape.anchor = (XSSFClientAnchor) anchor;

            return shape;
        }

        /**
         * Returns all charts in this Drawing.
         */
        public List<XSSFChart> GetCharts()
        {
            List<XSSFChart> charts = new List<XSSFChart>();
            foreach (POIXMLDocumentPart part in GetRelations())
            {
                if (part is XSSFChart chart)
                {
                    charts.Add(chart);
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

        private CT_Transform2D CreateXfrm(XSSFClientAnchor anchor)
        {
            CT_Transform2D xfrm = new CT_Transform2D();
            CT_Point2D off = xfrm.AddNewOff();
            off.x = anchor.Dx1;
            off.y = anchor.Dy1;
            XSSFSheet sheet = Sheet;
            double widthPx = 0;
            for(int col = anchor.Col1; col<anchor.Col2; col++)
            {
                widthPx += sheet.GetColumnWidthInPixels(col);
            }
            double heightPx = 0;
            for(int row = anchor.Row1; row<anchor.Row2; row++)
            {
                heightPx += ImageUtils.GetRowHeightInPixels(sheet, row);
            }
            long width = Units.PixelToEMU((int)widthPx);
            long height = Units.PixelToEMU((int)heightPx);
            CT_PositiveSize2D ext = xfrm.AddNewExt();
            ext.cx = (width - anchor.Dx1 + anchor.Dx2);
            ext.cy = (height - anchor.Dy1 + anchor.Dy2);

            // TODO: handle vflip/hflip
            return xfrm;
        }

        private long NewShapeId()
        {
            return drawing.SizeOfTwoCellAnchorArray() + 1;
        }

        #region IDrawing Members

        public bool ContainsChart()
        {
            throw new NotImplementedException();
        }

        #endregion

        /// <summary>
        /// get shapes in this shape group
        /// </summary>
        /// <param name="groupshape"></param>
        /// <returns>list of shapes in this shape group</returns>
        public List<XSSFShape> GetShapes(XSSFShapeGroup groupshape)
        {
            List<XSSFShape> lst = new List<XSSFShape>();
            var gs = groupshape.GetCTGroupShape();
            AddShapes(gs, lst);
            return lst;
        }

        private void AddShapes(CT_GroupShape gs, List<XSSFShape> lst)
        {
            XSSFShape shape = null;
            foreach(var sp in gs.Shapes)
            {
                shape = HasOleLink(sp)
                        ? new XSSFObjectData(this, sp)
                        : new XSSFSimpleShape(this, sp);
                shape.anchor = GetAnchorFromParent(sp.Node);
                lst.Add(shape);
            }
            foreach(var p in gs.Pictures)
            {
                shape = new XSSFPicture(this, p);
                shape.anchor = GetAnchorFromParent(p.Node);
                lst.Add(shape);
            }
            foreach (var c in gs.Connectors)
            {
                shape = new XSSFConnector(this, c);
                shape.anchor = GetAnchorFromParent(c.Node);
                lst.Add(shape);
            }
            foreach(var g in gs.Groups)
            {
                shape = new XSSFShapeGroup(this, g);
                shape.anchor = GetAnchorFromParent(g.Node);
                lst.Add(shape);
            }
        }
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
                else if(anchor.sp != null)
                {
                    shape = XSSFDrawing.HasOleLink(anchor.sp)
                        ? new XSSFObjectData(this, anchor.sp)
                        : new XSSFSimpleShape(this, anchor.sp);
                }
                else if (anchor.groupShape != null)
                {
                    shape = new XSSFShapeGroup(this, anchor.groupShape);
                    //List<object> lstCtShapes = new List<object>();
                    //anchor.groupShape.GetShapes(lstCtShapes);

                    //foreach(var s in lstCtShapes)
                    //{
                    //    XSSFShape gShape = null;
                    //    if(s is CT_Connector)
                    //    {
                    //        gShape = new XSSFConnector(this, (CT_Connector)s);
                    //    }
                    //    else if(s is CT_Picture)
                    //    {
                    //        gShape = new XSSFPicture(this, (CT_Picture)s);
                    //    }
                    //    else if(s is CT_Shape)
                    //    {
                    //        gShape = new XSSFSimpleShape(this, (CT_Shape)s);
                    //    }
                    //    else if(s is CT_GroupShape)
                    //    {
                    //        gShape = new XSSFShapeGroup(this, (CT_GroupShape)s);
                    //    }
                    //    else if(s is CT_GraphicalObjectFrame)
                    //    {
                    //        gShape = new XSSFGraphicFrame(this, (CT_GraphicalObjectFrame)s);
                    //    }
                    //    if(gShape != null)
                    //    {
                    //        gShape.anchor = GetAnchorFromIEGAnchor(anchor);
                    //        gShape.cellanchor = anchor;
                    //        lst.Add(gShape);
                    //    }
                    //}
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
                    shape.cellanchor = anchor;
                    lst.Add(shape);
                }
            }
            
            return lst;
        }

        private static bool HasOleLink(CT_Shape shape)
        {
            try
            {
                string uri = shape.nvSpPr.cNvPr.extLst.ext[0].uri;
                if("{63B3BB69-23CF-44E3-9099-C40C66FF867C}".Equals(uri))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
            return false;
        }
 
        private XSSFAnchor GetAnchorFromIEGAnchor(IEG_Anchor ctAnchor)
        {
            XSSFAnchor anchor = null;
            if (ctAnchor is CT_TwoCellAnchor cellAnchor)
            {
                CT_TwoCellAnchor ct = (CT_TwoCellAnchor)ctAnchor;
                anchor = new XSSFClientAnchor(ct.from, ct.to);
            }
            else if (ctAnchor is CT_OneCellAnchor oneCellAnchor)
            {
                CT_OneCellAnchor ct = (CT_OneCellAnchor)ctAnchor;
                anchor = new XSSFClientAnchor(Sheet, ct.from, ct.ext);
            }
            else if (ctAnchor is CT_AbsoluteCellAnchor)
            {
                CT_AbsoluteCellAnchor ct = (CT_AbsoluteCellAnchor)ctAnchor;
                anchor = new XSSFClientAnchor(Sheet, ct.pos, ct.ext);
            }
            return anchor;
        }
        public XSSFSheet Sheet
        {
            get
            {
                return (XSSFSheet)GetParent();
            }
            
        }

        private static XSSFAnchor GetAnchorFromParent(XmlNode obj)
        {
            XSSFAnchor anchor = null;
            XmlNode parentNode = obj.ParentNode;
            while (parentNode != null)
            {
                if(parentNode.LocalName=="twoCellAnchor"||parentNode.LocalName=="oneCellAnchor" ||
                    parentNode.LocalName=="absoluteAnchor")
                    break;
                parentNode = parentNode.ParentNode;
            }
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

        public IEnumerator<IShape> GetEnumerator()
        {
            return GetShapes().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetShapes().GetEnumerator();
        }
    }
}

