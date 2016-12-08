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
using System.Drawing;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.Util;
using System.Xml;
using NPOI.SS.Util;

namespace NPOI.XSSF.UserModel
{

    /**
     * Represents a picture shape in a SpreadsheetML Drawing.
     *
     * @author Yegor Kozlov
     */
    public class XSSFPicture : XSSFShape, IPicture
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(XSSFPicture));

        /**
         * Column width measured as the number of characters of the maximum digit width of the
         * numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. There are 4 pixels of margin
         * pAdding (two on each side), plus 1 pixel pAdding for the gridlines.
         *
         * This value is the same for default font in Office 2007 (Calibry) and Office 2003 and earlier (Arial)
         */
        //private static float DEFAULT_COLUMN_WIDTH = 9.140625f;

        /**
         * A default instance of CTShape used for creating new shapes.
         */
        private static CT_Picture prototype = null;

        /**
         * This object specifies a picture object and all its properties
         */
        private CT_Picture ctPicture;

        /**
         * Construct a new XSSFPicture object. This constructor is called from
         *  {@link XSSFDrawing#CreatePicture(XSSFClientAnchor, int)}
         *
         * @param Drawing the XSSFDrawing that owns this picture
         */
        public XSSFPicture(XSSFDrawing drawing, CT_Picture ctPicture)
        {
            this.drawing = drawing;
            this.ctPicture = ctPicture;
        }

        /**
         * Returns a prototype that is used to construct new shapes
         *
         * @return a prototype that is used to construct new shapes
         */
        public XSSFPicture(XSSFDrawing drawing, XmlNode ctPicture)
        {
            this.drawing = drawing;
            this.ctPicture =CT_Picture.Parse(ctPicture, POIXMLDocumentPart.NamespaceManager);
        }

        internal static CT_Picture Prototype()
        {

                CT_Picture pic = new CT_Picture();
                CT_PictureNonVisual nvpr = pic.AddNewNvPicPr();
                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualDrawingProps nvProps = nvpr.AddNewCNvPr();
                nvProps.id = (1);
                nvProps.name = ("Picture 1");
                nvProps.descr = ("Picture");
                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualPictureProperties nvPicProps = nvpr.AddNewCNvPicPr();
                nvPicProps.AddNewPicLocks().noChangeAspect = true;



                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_BlipFillProperties blip = pic.AddNewBlipFill();
                blip.AddNewBlip().embed = "";
                blip.AddNewStretch().AddNewFillRect();

                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties sppr = pic.AddNewSpPr();
                CT_Transform2D t2d = sppr.AddNewXfrm();
                CT_PositiveSize2D ext = t2d.AddNewExt();
                //should be original picture width and height expressed in EMUs
                ext.cx = (0);
                ext.cy = (0);

                CT_Point2D off = t2d.AddNewOff();
                off.x=(0);
                off.y=(0);

                CT_PresetGeometry2D prstGeom = sppr.AddNewPrstGeom();
                prstGeom.prst = (ST_ShapeType.rect);
                prstGeom.AddNewAvLst();

                prototype = pic;
            return prototype;
        }

        /**
         * Link this shape with the picture data
         *
         * @param rel relationship referring the picture data
         */
        internal void SetPictureReference(PackageRelationship rel)
        {
            ctPicture.blipFill.blip.embed = rel.Id;
        }

        /**
         * Return the underlying CT_Picture bean that holds all properties for this picture
         *
         * @return the underlying CT_Picture bean
         */

        internal CT_Picture GetCTPicture()
        {
            return ctPicture;
        }

        /**
         * Reset the image to the original size.
         *
         * <p>
         * Please note, that this method works correctly only for workbooks
         * with the default font size (Calibri 11pt for .xlsx).
         * If the default font is Changed the resized image can be streched vertically or horizontally.
         * </p>
         */
        public void Resize()
        {
            Resize(double.MaxValue);
        }
        /**
         * Resize the image proportionally.
         *
         * @see #resize(double, double)
         */
        public void Resize(double scale)
        {
            Resize(scale, scale);
        }
        /**
         * Resize the image relatively to its current size.
         * <p>
         * Please note, that this method works correctly only for workbooks
         * with the default font size (Calibri 11pt for .xlsx).
         * If the default font is changed the resized image can be streched vertically or horizontally.
         * </p>
         * <p>
         * <code>resize(1.0,1.0)</code> keeps the original size,<br/>
         * <code>resize(0.5,0.5)</code> resize to 50% of the original,<br/>
         * <code>resize(2.0,2.0)</code> resizes to 200% of the original.<br/>
         * <code>resize({@link Double#MAX_VALUE},{@link Double#MAX_VALUE})</code> resizes to the dimension of the embedded image. 
         * </p>
         *
         * @param scaleX the amount by which the image width is multiplied relative to the original width,
         *  when set to {@link java.lang.Double#MAX_VALUE} the width of the embedded image is used
         * @param scaleY the amount by which the image height is multiplied relative to the original height,
         *  when set to {@link java.lang.Double#MAX_VALUE} the height of the embedded image is used
         */
        public void Resize(double scaleX, double scaleY)
        {
            IClientAnchor anchor = (XSSFClientAnchor)GetAnchor();

            IClientAnchor pref = GetPreferredSize(scaleX, scaleY);

            int row2 = anchor.Row1 + (pref.Row2 - pref.Row1);
            int col2 = anchor.Col1 + (pref.Col2 - pref.Col1);

            anchor.Col2=(col2);
            //anchor.Dx1=(0);
            anchor.Dx2=(pref.Dx2);

            anchor.Row2=(row2);
            //anchor.Dy1=(0);
            anchor.Dy2=(pref.Dy2);
        }

        /**
         * Calculate the preferred size for this picture.
         *
         * @return XSSFClientAnchor with the preferred size for this image
         */
        public IClientAnchor GetPreferredSize()
        {
            return GetPreferredSize(1);
        }

        /**
         * Calculate the preferred size for this picture.
         *
         * @param scale the amount by which image dimensions are multiplied relative to the original size.
         * @return XSSFClientAnchor with the preferred size for this image
         */
        public IClientAnchor GetPreferredSize(double scale)
        {
            return GetPreferredSize(scale, scale);
        }


        /**
         * Calculate the preferred size for this picture.
         *
         * @param scaleX the amount by which image width is multiplied relative to the original width.
         * @param scaleY the amount by which image height is multiplied relative to the original height.
         * @return XSSFClientAnchor with the preferred size for this image
         */
        public IClientAnchor GetPreferredSize(double scaleX, double scaleY)
        {
            Size dim = ImageUtils.SetPreferredSize(this, scaleX, scaleY);
            CT_PositiveSize2D size2d = ctPicture.spPr.xfrm.ext;
            size2d.cx = (dim.Width);
            size2d.cy = (dim.Height);
            return ClientAnchor;
        }
        /**
         * Return the dimension of this image
         *
         * @param part the namespace part holding raw picture data
         * @param type type of the picture: {@link Workbook#PICTURE_TYPE_JPEG},
         * {@link Workbook#PICTURE_TYPE_PNG} or {@link Workbook#PICTURE_TYPE_DIB}
         *
         * @return image dimension in pixels
         */
        protected static Size GetImageDimension(PackagePart part, PictureType type)
        {
            try
            {
                //return Image.FromStream(part.GetInputStream()).Size;
                //java can only read png,jpeg,dib image
                //C# read the image that format defined by PictureType , maybe.
                return ImageUtils.GetImageDimension(part.GetInputStream());
            }
            catch (IOException e)
            {
                //return a "singulariry" if ImageIO failed to read the image
                logger.Log(POILogger.WARN, e);
                return new Size();
            }
        }
        /**
         * Return the dimension of the embedded image in pixel
         *
         * @return image dimension in pixels
         */
        public Size GetImageDimension()
        {
            XSSFPictureData picData = PictureData as XSSFPictureData;
            return GetImageDimension(picData.GetPackagePart(), picData.PictureType);
        }

        protected internal override NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties GetShapeProperties()
        {
            return ctPicture.spPr;
        }


        #region IShape Members

        public int CountOfAllChildren
        {
            get { throw new NotImplementedException(); }
        }

        public int FillColor
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public LineStyle LineStyle
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                base.LineStyle = value;
            }
        }

        public int LineStyleColor
        {
            get { throw new NotImplementedException(); }
        }

        public int LineWidth
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                base.LineWidth = (value);
            }
        }

        public void SetLineStyleColor(int lineStyleColor)
        {
            throw new NotImplementedException();
        }

        #endregion


        public IPictureData PictureData
        {
            get
            {
                String blipId = ctPicture.blipFill.blip.embed;
                return (XSSFPictureData)GetDrawing().GetRelationById(blipId);
            }
        }

        /**
         * @return the anchor that is used by this shape.
         */

        public IClientAnchor ClientAnchor
        {
            get
            {
                XSSFAnchor a = GetAnchor() as XSSFAnchor;
                return (a is XSSFClientAnchor) ? (XSSFClientAnchor)a : null;
            }
        }

        /**
         * @return the sheet which contains the picture shape
         */

        public ISheet Sheet
        {
            get
            {
                return (XSSFSheet)this.GetDrawing().GetParent();
            }
        }
    }
}

