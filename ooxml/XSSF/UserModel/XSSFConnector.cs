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

using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System.Xml;

namespace NPOI.XSSF.UserModel
{

    /**
     * A connection shape Drawing element. A connection shape is a line, etc.
     * that connects two other shapes in this Drawing.
     *
     * @author Yegor Kozlov
     */
    public class XSSFConnector : XSSFShape
    {

        private static CT_Connector prototype = null;

        private CT_Connector ctShape;

        /**
         * Construct a new XSSFConnector object.
         *
         * @param Drawing the XSSFDrawing that owns this shape
         * @param ctShape the shape bean that holds all the shape properties
         */
        public XSSFConnector(XSSFDrawing drawing, CT_Connector ctShape)
        {
            this.drawing = drawing;
            this.ctShape = ctShape;
        }
        /**
         * Initialize default structure of a new auto-shape
         *
         */
        public static CT_Connector Prototype()
        {

                CT_Connector shape = new CT_Connector();
                CT_ConnectorNonVisual nv = shape.AddNewNvCxnSpPr();
                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualDrawingProps nvp = nv.AddNewCNvPr();
                nvp.id = (1);
                nvp.name = ("Shape 1");
                nv.AddNewCNvCxnSpPr();

                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties sp = shape.AddNewSpPr();
                CT_Transform2D t2d = sp.AddNewXfrm();
                CT_PositiveSize2D p1 = t2d.AddNewExt();
                p1.cx = (0);
                p1.cy = (0);
                CT_Point2D p2 = t2d.AddNewOff();
                p2.x =(0);
                p2.y=(0);

                CT_PresetGeometry2D geom = sp.AddNewPrstGeom();
                geom.prst = (ST_ShapeType.line);
                geom.AddNewAvLst();

                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeStyle style = shape.AddNewStyle();
                CT_SchemeColor scheme = style.AddNewLnRef().AddNewSchemeClr();
                scheme.val = (ST_SchemeColorVal.accent1);
                style.lnRef.idx = (1);

                CT_StyleMatrixReference fillref = style.AddNewFillRef();
                fillref.idx = (0);
                fillref.AddNewSchemeClr().val=(ST_SchemeColorVal.accent1);

                CT_StyleMatrixReference effectRef = style.AddNewEffectRef();
                effectRef.idx = (0);
                effectRef.AddNewSchemeClr().val = (ST_SchemeColorVal.accent1);

                CT_FontReference fontRef = style.AddNewFontRef();
                fontRef.idx = (ST_FontCollectionIndex.minor);
                fontRef.AddNewSchemeClr().val = (ST_SchemeColorVal.tx1);

                prototype = shape;

            return prototype;
        }


        public CT_Connector GetCTConnector()
        {
            return ctShape;
        }

        /**
         * Gets the shape type, one of the constants defined in {@link NPOI.ss.usermodel.ShapeTypes}.
         *
         * @return the shape type
         * @see NPOI.ss.usermodel.ShapeTypes
         */
        public ST_ShapeType ShapeType
        {
            get
            {
                return ctShape.spPr.prstGeom.prst;
            }
            set 
            {
                ctShape.spPr.prstGeom.prst = value;
            }
        }
        protected internal override NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties GetShapeProperties()
        {
            return ctShape.spPr;
        }

    }
}



