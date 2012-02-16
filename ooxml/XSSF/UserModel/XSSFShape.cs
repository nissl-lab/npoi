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
namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a shape in a SpreadsheetML Drawing.
     *
     * @author Yegor Kozlov
     */
    public abstract class XSSFShape
    {
        public static int EMU_PER_PIXEL = 9525;
        public static int EMU_PER_POINT = 12700;

        public static int POINT_DPI = 72;
        public static int PIXEL_DPI = 96;

        /**
         * Parent Drawing
         */
        protected XSSFDrawing drawing;

        /**
         * The parent shape, always not-null for shapes in groups
         */
        public XSSFShapeGroup parent;

        /**
         * anchor that is used by this shape
         */
        internal XSSFAnchor anchor;

        /**
         * Return the Drawing that owns this shape
         *
         * @return the parent Drawing that owns this shape
         */
        public XSSFDrawing GetDrawing()
        {
            return drawing;
        }

        /**
         * Gets the parent shape.
         */
        public XSSFShapeGroup Parent
        {
            get
            {
                return parent;
            }
        }

        /**
         * @return  the anchor that is used by this shape.
         */
        public XSSFAnchor GetAnchor()
        {
            return anchor;
        }

        /**
         * Returns xml bean with shape properties.
         *
         * @return xml bean with shape properties.
         */
        protected abstract CT_ShapeProperties GetShapeProperties();

        /**
         * Whether this shape is not Filled with a color
         *
         * @return true if this shape is not Filled with a color.
         */
        public bool IsNoFill()
        {
            return GetShapeProperties().noFill!=null;
        }

        /**
         * Sets whether this shape is Filled or transparent.
         *
         * @param noFill if true then no fill will be applied to the shape element.
         */
        public void SetNoFill(bool noFill)
        {
            CT_ShapeProperties props = GetShapeProperties();
            //unset solid and pattern Fills if they are Set
            if (props.IsSetPattFill()) props.unsetPattFill();
            if (props.IsSetSolidFill()) props.unsetSolidFill();

            props.noFill = new CT_NoFillProperties();
        }

        /**
         * Sets the color used to fill this shape using the solid fill pattern.
         */
        public void SetFillColor(int red, int green, int blue)
        {
            CT_ShapeProperties props = GetShapeProperties();
            CT_SolidColorFillProperties fill = props.IsSetSolidFill() ? props.solidFill : props.AddNewSolidFill();
            CT_SRgbColor rgb = new CT_SRgbColor();
            rgb.val = (new byte[] { (byte)red, (byte)green, (byte)blue });
            fill.srgbClr = (rgb);
        }

        /**
         * The color applied to the lines of this shape.
         */
        public void SetLineStyleColor(int red, int green, int blue)
        {
            CT_ShapeProperties props = GetShapeProperties();
            CT_LineProperties ln = props.IsSetLn() ? props.ln : props.AddNewLn();
            CT_SolidColorFillProperties fill = ln.IsSetSolidFill() ? ln.solidFill : ln.AddNewSolidFill();
            CT_SRgbColor rgb = new CT_SRgbColor();
            rgb.val = (new byte[] { (byte)red, (byte)green, (byte)blue });
            fill.srgbClr = (rgb);
        }

        /**
         * Specifies the width to be used for the underline stroke.
         *
         * @param lineWidth width in points
         */
        public void SetLineWidth(double lineWidth)
        {
            CT_ShapeProperties props = GetShapeProperties();
            CT_LineProperties ln = props.IsSetLn() ? props.ln : props.AddNewLn();
            ln.w =(int)(lineWidth * EMU_PER_POINT);
        }

        /**
         * Sets the line style.
         *
         * @param lineStyle
         */
        public void SetLineStyle(int lineStyle)
        {
            CT_ShapeProperties props = GetShapeProperties();
            CT_LineProperties ln = props.IsSetLn() ? props.ln : props.AddNewLn();
            CT_PresetLineDashProperties dashStyle = CT_PresetLineDashProperties.Factory.newInstance();
            dashStyle.val = (ST_PresetLineDashVal)(lineStyle + 1);
            ln.prstDash = dashStyle;
        }

    }
}



