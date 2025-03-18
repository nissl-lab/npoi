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
using System.Collections.Generic;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.Util;

namespace NPOI.XSSF.UserModel
{
    /**
     * This object specifies a group shape that represents many shapes grouped together. This shape is to be treated
     * just as if it were a regular shape but instead of being described by a single geometry it is made up of all the
     * shape geometries encompassed within it. Within a group shape each of the shapes that make up the group are
     * specified just as they normally would.
     *
     * @author Yegor Kozlov
     */
    public class XSSFShapeGroup : XSSFShape
    {
        private static CT_GroupShape prototype = null;

        private CT_GroupShape ctGroup;

        /**
         * Construct a new XSSFSimpleShape object.
         *
         * @param Drawing the XSSFDrawing that owns this shape
         * @param ctGroup the XML bean that stores this group content
         */
        public XSSFShapeGroup(XSSFDrawing drawing, CT_GroupShape ctGroup)
        {
            this.drawing = drawing;
            this.ctGroup = ctGroup;
        }

        /**
         * Initialize default structure of a new shape group
         */
        internal static CT_GroupShape Prototype()
        {

                CT_GroupShape shape = new CT_GroupShape();


                CT_GroupShapeNonVisual nv = shape.AddNewNvGrpSpPr();
                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualDrawingProps nvpr = nv.AddNewCNvPr();
                nvpr.id = (0);
                nvpr.name = ("Group 0");
                nv.AddNewCNvGrpSpPr();
                CT_GroupShapeProperties sp = shape.AddNewGrpSpPr();
                CT_GroupTransform2D t2d = sp.AddNewXfrm();
                CT_PositiveSize2D p1 = t2d.AddNewExt();
                p1.cx = (0);
                p1.cy = (0);
                CT_Point2D p2 = t2d.AddNewOff();
                p2.x = (0);
                p2.y = (0);
                CT_PositiveSize2D p3 = t2d.AddNewChExt();
                p3.cx = (0);
                p3.cy = (0);
                CT_Point2D p4 = t2d.AddNewChOff();
                p4.x = (0);
                p4.y = (0);

                prototype = shape;
            return prototype;
        }

        /**
         * Returns the shape group id.
         * @return id of the shape group
         */
        public override uint ID
        {
            get
            {
                return ctGroup.nvGrpSpPr.cNvPr.id;
            }
        }

        /**
         * Returns the shape group name.
         * @return name of the shape group
         */
        public override String Name
        {
            get
            {
                return ctGroup.nvGrpSpPr.cNvPr.name;
            }
            set
            {
                ctGroup.nvGrpSpPr.cNvPr.name = value;
            }
        }

        /**
         * Constructs a textbox.
         *
         * @param anchor the child anchor describes how this shape is attached
         *               to the group.
         * @return      the newly Created textbox.
         */
        public XSSFTextBox CreateTextbox(XSSFChildAnchor anchor)
        {
            CT_Shape ctShape = ctGroup.AddNewSp();
            ctShape.Set(XSSFSimpleShape.Prototype());

            XSSFTextBox shape = new XSSFTextBox(GetDrawing(), ctShape);
            shape.parent = this;
            shape.anchor = anchor;
            shape.GetCTShape().spPr.xfrm = (anchor.GetCTTransform2D());
            return shape;

        }
        /**
         * Creates a simple shape.  This includes such shapes as lines, rectangles,
         * and ovals.
         *
         * @param anchor the child anchor describes how this shape is attached
         *               to the group.
         * @return the newly Created shape.
         */
        public XSSFSimpleShape CreateSimpleShape(XSSFChildAnchor anchor)
        {
            CT_Shape ctShape = ctGroup.AddNewSp();
            ctShape.Set(XSSFSimpleShape.Prototype());

            XSSFSimpleShape shape = new XSSFSimpleShape(GetDrawing(), ctShape);
            shape.parent = (this);
            shape.anchor = anchor;
            shape.GetCTShape().spPr.xfrm = (anchor.GetCTTransform2D());
            return shape;
        }

        /**
         * Creates a simple shape.  This includes such shapes as lines, rectangles,
         * and ovals.
         *
         * @param anchor the child anchor describes how this shape is attached
         *               to the group.
         * @return the newly Created shape.
         */
        public XSSFConnector CreateConnector(XSSFChildAnchor anchor)
        {
            CT_Connector ctShape = ctGroup.AddNewCxnSp();
            ctShape.Set(XSSFConnector.Prototype());

            XSSFConnector shape = new XSSFConnector(GetDrawing(), ctShape);
            shape.parent = this;
            shape.anchor = anchor;
            shape.GetCTConnector().spPr.xfrm = (anchor.GetCTTransform2D());
            return shape;
        }

        /**
         * Creates a picture.
         *
         * @param anchor       the client anchor describes how this picture is attached to the sheet.
         * @param pictureIndex the index of the picture in the workbook collection of pictures,
         *                     {@link XSSFWorkbook#getAllPictures()} .
         * @return the newly Created picture shape.
         */
        public XSSFPicture CreatePicture(XSSFChildAnchor anchor, int pictureIndex)
        {
            PackageRelationship rel = GetDrawing().AddPictureReference(pictureIndex);

            CT_Picture ctShape = ctGroup.AddNewPic();
            ctShape.Set(XSSFPicture.Prototype());

            XSSFPicture shape = new XSSFPicture(GetDrawing(), ctShape);
            shape.parent = this;
            shape.anchor = anchor;
            shape.SetPictureReference(rel);
            shape.GetCTPicture().spPr.xfrm = (anchor.GetCTTransform2D());
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
              BuildFreeForm BFF
        ) {
            var anchor = new XSSFChildAnchor((int)BFF.Left, (int)BFF.Top
                                            , (int)BFF.Rigth, (int)BFF.Bottom);
            //long shapeId = newShapeId();
            CT_Shape ctShape = ctGroup.AddNewSp();
            ctShape.Set(XSSFFreeform.Prototype());
            //ctShape.nvSpPr.cNvPr.id = (uint) (shapeId);
            var freeform = new XSSFFreeform(GetDrawing(), ctShape);
            freeform.parent = this;
            freeform.anchor = anchor;

            freeform.Build(BFF);

            freeform.GetCTShape().spPr.xfrm = (anchor.GetCTTransform2D());

            return freeform;
        }

        public XSSFShapeGroup CreateGroup(XSSFChildGroupAnchor anchor)
        {
            CT_GroupShape ctShape = ctGroup.AddNewGroup();
            ctShape.Set(XSSFShapeGroup.Prototype());

            XSSFShapeGroup group = new XSSFShapeGroup(GetDrawing(), ctShape)
            {
                parent = this,
                anchor = anchor
            };
            group.GetCTGroupShape().grpSpPr.xfrm = anchor.GetCTTransform2D();

            return group;
        }

        public CT_GroupShape GetCTGroupShape()
        {
            return ctGroup;
        }

        /**
         * Sets the coordinate space of this group.  All children are constrained
         * to these coordinates.
         */
        public void SetCoordinates(int x1, int y1, int x2, int y2)
        {
            CT_GroupTransform2D t2d = ctGroup.grpSpPr.xfrm;
            CT_Point2D off = t2d.off;
            off.x = (x1);
            off.y = (y1);
            CT_PositiveSize2D ext = t2d.ext;
            ext.cx = (x2);
            ext.cy = (y2);

            CT_Point2D chOff = t2d.chOff;
            chOff.x = (x1);
            chOff.y = (y1);
            CT_PositiveSize2D chExt = t2d.chExt;
            chExt.cx = (x2);
            chExt.cy = (y2);
        }

        protected internal override NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties GetShapeProperties()
        {
            throw new InvalidOperationException("Not supported for shape group");
        }

        public void AutoFit(SS.UserModel.ISheet Sheet) {
            var min = new Coords(long.MaxValue, long.MaxValue);
            var max = new Coords(long.MinValue, long.MinValue);

            AutoFit(ctGroup, min, max);

            ctGroup.grpSpPr.xfrm.off.x      = min.x;
            ctGroup.grpSpPr.xfrm.off.y      = min.y;
            ctGroup.grpSpPr.xfrm.ext.cx     = max.x - min.x;
            ctGroup.grpSpPr.xfrm.ext.cy     = max.y - min.y;
            ctGroup.grpSpPr.xfrm.chOff.x    = ctGroup.grpSpPr.xfrm.off.x;
            ctGroup.grpSpPr.xfrm.chOff.y    = ctGroup.grpSpPr.xfrm.off.y;
            ctGroup.grpSpPr.xfrm.chExt.cx   = ctGroup.grpSpPr.xfrm.ext.cx;
            ctGroup.grpSpPr.xfrm.chExt.cy   = ctGroup.grpSpPr.xfrm.ext.cy;

            if(parent == null) { // top level group
                var ac = new XSSFClientAnchor(Sheet, (int)min.x, (int)min.y, (int)max.x, (int)max.y);
                if(cellanchor is CT_TwoCellAnchor cellAnchor) {
                    ((CT_TwoCellAnchor) cellanchor).from = ac.From;
                    ((CT_TwoCellAnchor) cellanchor).to= ac.To;
                }
            }
        }

        private void AutoFit(CT_GroupShape CtGroup, Coords Min, Coords Max) {
            foreach(var cxn in CtGroup.Connectors) {
                var cd = new Coords(cxn.spPr.xfrm.off.x, cxn.spPr.xfrm.off.y);
                Min.Min(cd);
                Max.Max(cd);
                cd += new Coords(cxn.spPr.xfrm.ext.cx, cxn.spPr.xfrm.ext.cy);
                Min.Min(cd);
                Max.Max(cd);
            }
            foreach(var pic in CtGroup.Pictures) {
                var cd = new Coords(pic.spPr.xfrm.off.x, pic.spPr.xfrm.off.y);
                Min.Min(cd);
                Max.Max(cd);
                cd += new Coords(pic.spPr.xfrm.ext.cx, pic.spPr.xfrm.ext.cy);
                Min.Min(cd);
                Max.Max(cd);
            }
            foreach(var sp in CtGroup.Shapes) {
                var cd = new Coords(sp.spPr.xfrm.off.x, sp.spPr.xfrm.off.y);
                Min.Min(cd);
                Max.Max(cd);
                cd += new Coords(sp.spPr.xfrm.ext.cx, sp.spPr.xfrm.ext.cy);
                Min.Min(cd);
                Max.Max(cd);
            }
            foreach(var gp in CtGroup.Groups) {
                var min = new Coords(long.MaxValue, long.MaxValue);
                var max = new Coords(long.MinValue, long.MinValue);

                AutoFit(gp, min, max);

                gp.grpSpPr.xfrm.off.x = min.x;
                gp.grpSpPr.xfrm.off.y = min.y;
                gp.grpSpPr.xfrm.ext.cx = max.x - min.x;
                gp.grpSpPr.xfrm.ext.cy = max.y - min.y;
                gp.grpSpPr.xfrm.chOff.x = gp.grpSpPr.xfrm.off.x;
                gp.grpSpPr.xfrm.chOff.y = gp.grpSpPr.xfrm.off.y;
                gp.grpSpPr.xfrm.chExt.cx = gp.grpSpPr.xfrm.ext.cx;
                gp.grpSpPr.xfrm.chExt.cy = gp.grpSpPr.xfrm.ext.cy;

                Min.Min(min);
                Max.Max(max);
            }
        }
    }
}
