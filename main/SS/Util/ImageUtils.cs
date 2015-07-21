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
namespace NPOI.SS.Util
{
    using System;
    using System.Drawing;
    using System.IO;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    /**
     * @author Yegor Kozlov
     */
    public class ImageUtils
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(ImageUtils));

        public static int PIXEL_DPI = 96;


        public static Size GetImageDimension(Stream is1)
        {
            using (Image img = Image.FromStream(is1))
            {
                //return img.Size;
                int[] dpi = GetResolution(img);

                //if DPI is zero then assume standard 96 DPI
                //since cannot divide by zero
                if (dpi[0] == 0) dpi[0] = PIXEL_DPI;
                if (dpi[1] == 0) dpi[1] = PIXEL_DPI;
                Size size = new Size();
                size.Width = img.Width * PIXEL_DPI / dpi[0];
                size.Height = img.Height * PIXEL_DPI / dpi[1];
                return size;
            }
        }
        /**
         * Return the dimension of this image
         *
         * @param is the stream Containing the image data
         * @param type type of the picture: {@link NPOI.SS.UserModel.Workbook#PICTURE_TYPE_JPEG},
         * {@link NPOI.SS.UserModel.Workbook#PICTURE_TYPE_PNG} or {@link NPOI.SS.UserModel.Workbook#PICTURE_TYPE_DIB}
         *
         * @return image dimension in pixels
         */
        public static Size GetImageDimension(Stream is1, PictureType type)
        {
            Size size = new Size();

            switch (type)
            {
                case PictureType.JPEG:
                case PictureType.PNG:
                case PictureType.DIB:
                    //we can calculate the preferred size only for JPEG, PNG and BMP
                    //other formats like WMF, EMF and PICT are not supported in Java
                    using (Image img = Image.FromStream(is1))
                    {
                        int[] dpi = GetResolution(img);

                        //if DPI is zero then assume standard 96 DPI
                        //since cannot divide by zero
                        if (dpi[0] == 0) dpi[0] = PIXEL_DPI;
                        if (dpi[1] == 0) dpi[1] = PIXEL_DPI;

                        size.Width = img.Width * PIXEL_DPI / dpi[0];
                        size.Height = img.Height * PIXEL_DPI / dpi[1];
                        return size;
                    }
                    
                default:
                    logger.Log(POILogger.WARN, "Only JPEG, PNG and DIB pictures can be automatically sized");
                    break;
            }
            return size;
        }
        
    

        /**
         * The metadata of PNG and JPEG can contain the width of a pixel in millimeters.
         * Return the the "effective" dpi calculated as <code>25.4/HorizontalPixelSize</code>
         * and <code>25.4/VerticalPixelSize</code>.  Where 25.4 is the number of mm in inch.
         *
         * @return array of two elements: <code>{horisontalPdi, verticalDpi}</code>.
         * {96, 96} is the default.
         */
        public static int[] GetResolution(Image r)
        {
            return new int[] { (int)r.HorizontalResolution, (int)r.VerticalResolution };
        }

        /**
         * Calculate and Set the preferred size (anchor) for this picture.
         *
         * @param scaleX the amount by which image width is multiplied relative to the original width.
         * @param scaleY the amount by which image height is multiplied relative to the original height.
         * @return the new Dimensions of the scaled picture in EMUs
         */
        public static Size SetPreferredSize(IPicture picture, double scaleX, double scaleY)
        {
            IClientAnchor anchor = picture.ClientAnchor;
            bool isHSSF = (anchor is HSSFClientAnchor);
            IPictureData data = picture.PictureData;
            ISheet sheet = picture.Sheet;

            // in pixel
            Size imgSize = GetImageDimension(new MemoryStream(data.Data), data.PictureType);
            // in emus
            Size anchorSize = ImageUtils.GetDimensionFromAnchor(picture);
            double scaledWidth = (scaleX == Double.MaxValue)
                ? imgSize.Width : anchorSize.Width / Units.EMU_PER_PIXEL * scaleX;
            double scaledHeight = (scaleY == Double.MaxValue)
                ? imgSize.Height : anchorSize.Height / Units.EMU_PER_PIXEL * scaleY;

            double w = 0;
            int col2 = anchor.Col1;
            int dx2 = 0;

            //space in the leftmost cell
            w = sheet.GetColumnWidthInPixels(col2++);
            if (isHSSF)
            {
                w *= 1 - anchor.Dx1 / 1024d;
            }
            else
            {
                w -= anchor.Dx1 / Units.EMU_PER_PIXEL;
            }

            while (w < scaledWidth)
            {
                w += sheet.GetColumnWidthInPixels(col2++);
            }

            if (w > scaledWidth)
            {
                //calculate dx2, offset in the rightmost cell
                double cw = sheet.GetColumnWidthInPixels(--col2);
                double delta = w - scaledWidth;
                if (isHSSF)
                {
                    dx2 = (int)((cw - delta) / cw * 1024);
                }
                else
                {
                    dx2 = (int)((cw - delta) * Units.EMU_PER_PIXEL);
                }
                if (dx2 < 0) dx2 = 0;
            }
            anchor.Col2 = (/*setter*/col2);
            anchor.Dx2 = (/*setter*/dx2);

            double h = 0;
            int row2 = anchor.Row1;
            int dy2 = 0;

            h = GetRowHeightInPixels(sheet, row2++);
            if (isHSSF)
            {
                h *= 1 - anchor.Dy1 / 256d;
            }
            else
            {
                h -= anchor.Dy1 / Units.EMU_PER_PIXEL;
            }

            while (h < scaledHeight)
            {
                h += GetRowHeightInPixels(sheet, row2++);
            }

            if (h > scaledHeight)
            {
                double ch = GetRowHeightInPixels(sheet, --row2);
                double delta = h - scaledHeight;
                if (isHSSF)
                {
                    dy2 = (int)((ch - delta) / ch * 256);
                }
                else
                {
                    dy2 = (int)((ch - delta) * Units.EMU_PER_PIXEL);
                }
                if (dy2 < 0) dy2 = 0;
            }

            anchor.Row2 = (/*setter*/row2);
            anchor.Dy2 = (/*setter*/dy2);

            Size dim = new Size(
                (int)Math.Round(scaledWidth * Units.EMU_PER_PIXEL),
                (int)Math.Round(scaledHeight * Units.EMU_PER_PIXEL)
            );

            return dim;
        }

        /**
         * Calculates the dimensions in EMUs for the anchor of the given picture
         *
         * @param picture the picture Containing the anchor
         * @return the dimensions in EMUs
         */
        public static Size GetDimensionFromAnchor(IPicture picture)
        {
            IClientAnchor anchor = picture.ClientAnchor;
            bool isHSSF = (anchor is HSSFClientAnchor);
            ISheet sheet = picture.Sheet;

            double w = 0;
            int col2 = anchor.Col1;

            //space in the leftmost cell
            w = sheet.GetColumnWidthInPixels(col2++);
            if (isHSSF)
            {
                w *= 1 - anchor.Dx1 / 1024d;
            }
            else
            {
                w -= anchor.Dx1 / Units.EMU_PER_PIXEL;
            }

            while (col2 < anchor.Col2)
            {
                w += sheet.GetColumnWidthInPixels(col2++);
            }

            if (isHSSF)
            {
                w += sheet.GetColumnWidthInPixels(col2) * anchor.Dx2 / 1024d;
            }
            else
            {
                w += anchor.Dx2 / Units.EMU_PER_PIXEL;
            }

            double h = 0;
            int row2 = anchor.Row1;

            h = GetRowHeightInPixels(sheet, row2++);
            if (isHSSF)
            {
                h *= 1 - anchor.Dy1 / 256d;
            }
            else
            {
                h -= anchor.Dy1 / Units.EMU_PER_PIXEL;
            }

            while (row2 < anchor.Row2)
            {
                h += GetRowHeightInPixels(sheet, row2++);
            }

            if (isHSSF)
            {
                h += GetRowHeightInPixels(sheet, row2) * anchor.Dy2 / 256;
            }
            else
            {
                h += anchor.Dy2 / Units.EMU_PER_PIXEL;
            }

            return new Size((int)w * Units.EMU_PER_PIXEL, (int)h * Units.EMU_PER_PIXEL);
        }


        private static double GetRowHeightInPixels(ISheet sheet, int rowNum)
        {
            IRow r = sheet.GetRow(rowNum);
            double points = (r == null) ? sheet.DefaultRowHeightInPoints : r.HeightInPoints;
            return Units.ToEMU(points) / Units.EMU_PER_PIXEL;
        }
    }

}