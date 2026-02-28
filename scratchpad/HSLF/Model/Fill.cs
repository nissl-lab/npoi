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

namespace NPOI.HSLF.Model;

using NPOI.ddf.*;
using NPOI.HSLF.record.*;
using NPOI.HSLF.usermodel.PictureData;
using NPOI.HSLF.usermodel.SlideShow;
using NPOI.util.POILogger;
using NPOI.util.POILogFactory;



/**
 * Represents functionality provided by the 'Fill Effects' dialog in PowerPoint.
 *
 * @author Yegor Kozlov
 */
public class Fill {
    // For logging
    protected POILogger logger = POILogFactory.GetLogger(this.GetClass());

    /**
     *  Fill with a solid color
     */
    public static int FILL_SOLID = 0;

    /**
     *  Fill with a pattern (bitmap)
     */
    public static int FILL_PATTERN = 1;

    /**
     *  A texture (pattern with its own color map)
     */
    public static int FILL_TEXTURE = 2;

    /**
     *  Center a picture in the shape
     */
    public static int FILL_PICTURE = 3;

    /**
     *  Shade from start to end points
     */
    public static int FILL_SHADE = 4;

    /**
     *  Shade from bounding rectangle to end point
     */
    public static int FILL_SHADE_CENTER = 5;

    /**
     *  Shade from shape outline to end point
     */
    public static int FILL_SHADE_SHAPE = 6;

    /**
     *  Similar to FILL_SHADE, but the fill angle
     *  is Additionally scaled by the aspect ratio of
     *  the shape. If shape is square, it is the same as FILL_SHADE
     */
    public static int FILL_SHADE_SCALE = 7;

    /**
     *  shade to title
     */
    public static int FILL_SHADE_TITLE = 8;

    /**
     *  Use the background fill color/pattern
     */
    public static int FILL_BACKGROUND = 9;



    /**
     * The shape this background applies to
     */
    protected Shape shape;

    /**
     * Construct a <code>Fill</code> object for a shape.
     * Fill information will be read from shape's escher properties.
     *
     * @param shape the shape this background applies to
     */
    public Fill(Shape shape){
        this.shape = shape;
    }

    /**
     * Returns fill type.
     * Must be one of the <code>FILL_*</code> constants defined in this class.
     *
     * @return type of fill
     */
    public int GetFillType(){
        EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(shape.GetSpContainer(), EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)Shape.GetEscherProperty(opt, EscherProperties.FILL__FILLTYPE);
        return prop == null ? FILL_SOLID : prop.GetPropertyValue();
    }

    /**
     * Sets fill type.
     * Must be one of the <code>FILL_*</code> constants defined in this class.
     *
     * @param type type of the fill
     */
    public void SetFillType(int type){
        EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(shape.GetSpContainer(), EscherOptRecord.RECORD_ID);
        Shape.SetEscherProperty(opt, EscherProperties.FILL__FILLTYPE, type);
    }

    /**
     * Foreground color
     */
    public Color GetForegroundColor(){
        EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(shape.GetSpContainer(), EscherOptRecord.RECORD_ID);
        EscherSimpleProperty p1 = (EscherSimpleProperty)Shape.GetEscherProperty(opt, EscherProperties.FILL__FILLCOLOR);
        EscherSimpleProperty p2 = (EscherSimpleProperty)Shape.GetEscherProperty(opt, EscherProperties.FILL__NOFILLHITTEST);
        EscherSimpleProperty p3 = (EscherSimpleProperty)Shape.GetEscherProperty(opt, EscherProperties.FILL__FILLOPACITY);

        int p2val = p2 == null ? 0 : p2.GetPropertyValue();
        int alpha =  p3 == null ? 255 : ((p3.GetPropertyValue() >> 8) & 0xFF);

        Color clr = null;
        if (p1 != null && (p2val  & 0x10) != 0){
            int rgb = p1.GetPropertyValue();
            clr = shape.GetColor(rgb, alpha);
        }
        return clr;
    }

    /**
     * Foreground color
     */
    public void SetForegroundColor(Color color){
        EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(shape.GetSpContainer(), EscherOptRecord.RECORD_ID);
        if (color == null) {
            Shape.SetEscherProperty(opt, EscherProperties.FILL__NOFILLHITTEST, 0x150000);
        }
        else {
            int rgb = new Color(color.GetBlue(), color.GetGreen(), color.GetRed(), 0).GetRGB();
            Shape.SetEscherProperty(opt, EscherProperties.FILL__FILLCOLOR, rgb);
            Shape.SetEscherProperty(opt, EscherProperties.FILL__NOFILLHITTEST, 0x150011);
        }
    }

    /**
     * Background color
     */
    public Color GetBackgroundColor(){
        EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(shape.GetSpContainer(), EscherOptRecord.RECORD_ID);
        EscherSimpleProperty p1 = (EscherSimpleProperty)Shape.GetEscherProperty(opt, EscherProperties.FILL__FILLBACKCOLOR);
        EscherSimpleProperty p2 = (EscherSimpleProperty)Shape.GetEscherProperty(opt, EscherProperties.FILL__NOFILLHITTEST);

        int p2val = p2 == null ? 0 : p2.GetPropertyValue();

        Color clr = null;
        if (p1 != null && (p2val  & 0x10) != 0){
            int rgb = p1.GetPropertyValue();
            clr = shape.GetColor(rgb, 255);
        }
        return clr;
    }

    /**
     * Background color
     */
    public void SetBackgroundColor(Color color){
        EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(shape.GetSpContainer(), EscherOptRecord.RECORD_ID);
        if (color == null) {
            Shape.SetEscherProperty(opt, EscherProperties.FILL__FILLBACKCOLOR, -1);
        }
        else {
            int rgb = new Color(color.GetBlue(), color.GetGreen(), color.GetRed(), 0).GetRGB();
            Shape.SetEscherProperty(opt, EscherProperties.FILL__FILLBACKCOLOR, rgb);
        }
    }

    /**
     * <code>PictureData</code> object used in a texture, pattern of picture Fill.
     */
    public PictureData GetPictureData(){
        EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(shape.GetSpContainer(), EscherOptRecord.RECORD_ID);
        EscherSimpleProperty p = (EscherSimpleProperty)Shape.GetEscherProperty(opt, EscherProperties.FILL__PATTERNTEXTURE);
        if (p == null) return null;

        SlideShow ppt = shape.Sheet.GetSlideShow();
        PictureData[] pict = ppt.GetPictureData();
        Document doc = ppt.GetDocumentRecord();

        EscherContainerRecord dggContainer = doc.GetPPDrawingGroup().GetDggContainer();
        EscherContainerRecord bstore = (EscherContainerRecord)Shape.GetEscherChild(dggContainer, EscherContainerRecord.BSTORE_CONTAINER);

        java.util.List<EscherRecord> lst = bstore.GetChildRecords();
        int idx = p.GetPropertyValue();
        if (idx == 0){
            logger.log(POILogger.WARN, "no reference to picture data found ");
        } else {
            EscherBSERecord bse = (EscherBSERecord)lst.Get(idx - 1);
            for ( int i = 0; i < pict.Length; i++ ) {
                if (pict[i].GetOffset() ==  bse.GetOffset()){
                    return pict[i];
                }
            }
        }

        return null;
    }

    /**
     * Assign picture used to fill the underlying shape.
     *
     * @param idx 0-based index of the picture Added to this ppt by <code>SlideShow.AddPicture</code> method.
     */
    public void SetPictureData(int idx){
        EscherOptRecord opt = (EscherOptRecord)Shape.GetEscherChild(shape.GetSpContainer(), EscherOptRecord.RECORD_ID);
        Shape.SetEscherProperty(opt, (short)(EscherProperties.FILL__PATTERNTEXTURE + 0x4000), idx);
    }

}





