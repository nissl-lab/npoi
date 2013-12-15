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
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.HSSF.Util;
using System.Xml;

namespace NPOI.XSSF.UserModel
{
    /**
     * Represents a shape with a predefined geometry in a SpreadsheetML Drawing.
     * Possible shape types are defined in {@link NPOI.ss.usermodel.ShapeTypes}
     *
     * @author Yegor Kozlov
     */
    public class XSSFSimpleShape : XSSFShape
    { // TODO - instantiable superclass
        /**
         * A default instance of CT_Shape used for creating new shapes.
         */
        private static CT_Shape prototype = null;

        /**
         *  Xml bean that stores properties of this shape
         */
        private CT_Shape ctShape;

        public XSSFSimpleShape(XSSFDrawing drawing, CT_Shape ctShape)
        {
            this.drawing = drawing;
            this.ctShape = ctShape;
        }


        public XSSFSimpleShape(XSSFDrawing drawing, XmlNode spNode)
        {
            this.drawing = drawing;
            this.ctShape = CT_Shape.Parse(spNode, null);
        }

        /**
         * Prototype with the default structure of a new auto-shape.
         */
        internal static CT_Shape Prototype()
        {
            if (prototype == null)
            {
                CT_Shape shape = new CT_Shape();


                CT_ShapeNonVisual nv = shape.AddNewNvSpPr();
                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualDrawingProps nvp = nv.AddNewCNvPr();
                nvp.id = (1);
                nvp.name = ("Shape 1");
                nv.AddNewCNvSpPr();

                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties sp = shape.AddNewSpPr();
                NPOI.OpenXmlFormats.Dml.CT_Transform2D t2d = sp.AddNewXfrm();
                CT_PositiveSize2D p1 = t2d.AddNewExt();
                p1.cx = (0);
                p1.cy = (0);
                CT_Point2D p2 = t2d.AddNewOff();
                p2.x = (0);
                p2.y = (0);

                CT_PresetGeometry2D geom = sp.AddNewPrstGeom();
                geom.prst = (ST_ShapeType.rect);
                geom.AddNewAvLst();

                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeStyle style = shape.AddNewStyle();
                CT_SchemeColor scheme = style.AddNewLnRef().AddNewSchemeClr();
                scheme.val = (ST_SchemeColorVal.accent1);
                scheme.AddNewShade().val = 50000;
                style.lnRef.idx = (2);

                CT_StyleMatrixReference Fillref = style.AddNewFillRef();
                Fillref.idx = (1);
                Fillref.AddNewSchemeClr().val = (ST_SchemeColorVal.accent1);

                CT_StyleMatrixReference effectRef = style.AddNewEffectRef();
                effectRef.idx = (0);
                effectRef.AddNewSchemeClr().val = (ST_SchemeColorVal.accent1);

                CT_FontReference fontRef = style.AddNewFontRef();
                fontRef.idx = (ST_FontCollectionIndex.minor);
                fontRef.AddNewSchemeClr().val = (ST_SchemeColorVal.lt1);

                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TextBody body = shape.AddNewTxBody();
                CT_TextBodyProperties bodypr = body.AddNewBodyPr();
                bodypr.anchor = (ST_TextAnchoringType.ctr);
                bodypr.rtlCol = (false);
                CT_TextParagraph p = body.AddNewP();

                p.AddNewPPr().algn = (ST_TextAlignType.ctr);
                CT_TextCharacterProperties endPr = p.AddNewEndParaRPr();
                endPr.lang = ("en-US");
                endPr.sz = (1100);

                body.AddNewLstStyle();

                prototype = shape;
            }
            return prototype;
        }


        public CT_Shape GetCTShape()
        {
            return ctShape;
        }

        /**
         * Gets the shape type, one of the constants defined in {@link NPOI.ss.usermodel.ShapeTypes}.
         *
         * @return the shape type
         * @see NPOI.ss.usermodel.ShapeTypes
         */
        public int GetShapeType()
        {
            return (int)ctShape.spPr.prstGeom.prst;
        }

        /**
         * Sets the shape types.
         *
         * @param type the shape type, one of the constants defined in {@link NPOI.ss.usermodel.ShapeTypes}.
         * @see NPOI.ss.usermodel.ShapeTypes
         */
        public void SetShapeType(int type)
        {
            ctShape.spPr.prstGeom.prst = ((ST_ShapeType)(type));
        }

        protected internal override NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties GetShapeProperties()
        {
            return ctShape.spPr;
        }

        public void SetText(XSSFRichTextString str)
        {

            XSSFWorkbook wb = (XSSFWorkbook)GetDrawing().GetParent().GetParent();
            str.SetStylesTableReference(wb.GetStylesSource());

            CT_TextParagraph p = new CT_TextParagraph();
            if (str.NumFormattingRuns == 0)
            {
                CT_RegularTextRun r = p.AddNewR();
                CT_TextCharacterProperties rPr = r.AddNewRPr();
                rPr.lang = ("en-US");
                rPr.sz = (1100);
                r.t = str.String;

            }
            else
            {
                for (int i = 0; i < str.GetCTRst().sizeOfRArray(); i++)
                {
                    CT_RElt lt = str.GetCTRst().GetRArray(i);
                    CT_RPrElt ltPr = lt.rPr;
                    if (ltPr == null) ltPr = lt.AddNewRPr();

                    CT_RegularTextRun r = p.AddNewR();
                    CT_TextCharacterProperties rPr = r.AddNewRPr();
                    rPr.lang = ("en-US");

                    ApplyAttributes(ltPr, rPr);

                    r.t = (lt.t);
                }
            }
            ctShape.txBody.SetPArray(new CT_TextParagraph[] { p });

        }

        /**
         *
         * org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRPrElt to
         * org.openxmlformats.schemas.drawingml.x2006.main.CTFont adapter
         */
        private static void ApplyAttributes(CT_RPrElt pr, CT_TextCharacterProperties rPr)
        {

            if (pr.sizeOfBArray() > 0) rPr.b = (pr.GetBArray(0).val);
            if (pr.sizeOfUArray() > 0)
            {
                ST_UnderlineValues u1 = pr.GetUArray(0).val;
                if (u1 == ST_UnderlineValues.single) rPr.u = (ST_TextUnderlineType.sng);
                else if (u1 == ST_UnderlineValues.@double) rPr.u = (ST_TextUnderlineType.dbl);
                else if (u1 == ST_UnderlineValues.none) rPr.u = (ST_TextUnderlineType.none);
            }
            if (pr.sizeOfIArray() > 0) rPr.i = (pr.GetIArray(0).val);

            if (pr.sizeOfFamilyArray() > 0)
            {
                CT_TextFont rFont = rPr.AddNewLatin();
                rFont.typeface = (pr.GetRFontArray(0).val);
            }

            if (pr.sizeOfSzArray() > 0)
            {
                int sz = (int)(pr.GetSzArray(0).val * 100);
                rPr.sz = (sz);
            }

            if (pr.sizeOfColorArray() > 0)
            {
                CT_SolidColorFillProperties fill = rPr.IsSetSolidFill() ? rPr.solidFill : rPr.AddNewSolidFill();
                NPOI.OpenXmlFormats.Spreadsheet.CT_Color xlsColor = pr.GetColorArray(0);
                if (xlsColor.IsSetRgb())
                {
                    CT_SRgbColor clr = fill.IsSetSrgbClr() ? fill.srgbClr : fill.AddNewSrgbClr();
                    clr.val = (xlsColor.rgb);
                }
                else if (xlsColor.IsSetIndexed())
                {
                    HSSFColor indexed = HSSFColor.GetIndexHash()[(int)xlsColor.indexed] as HSSFColor;
                    if (indexed != null)
                    {
                        byte[] rgb = new byte[3];
                        rgb[0] = (byte)indexed.GetTriplet()[0];
                        rgb[1] = (byte)indexed.GetTriplet()[1];
                        rgb[2] = (byte)indexed.GetTriplet()[2];
                        CT_SRgbColor clr = fill.IsSetSrgbClr() ? fill.srgbClr : fill.AddNewSrgbClr();
                        clr.val = (rgb);
                    }
                }
            }
        }
    }

}

