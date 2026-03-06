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
    using System.IO;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using SkiaSharp;
    
    /**
     * @author Yegor Kozlov
     */
    public class ImageUtils
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(ImageUtils));

        public static int PIXEL_DPI = 96;


        public static SKSizeI GetImageDimension(Stream is1)
        {
            byte[] data;
            using (var ms = new MemoryStream())
            {
                is1.CopyTo(ms);
                data = ms.ToArray();
            }
            using (SKBitmap img = SKBitmap.Decode(data))
            {
                if (img == null) return new SKSizeI();
                int[] dpi = GetResolutionFromBytes(data);

                //if DPI is zero then assume standard 96 DPI
                //since cannot divide by zero
                if (dpi[0] == 0) dpi[0] = PIXEL_DPI;
                if (dpi[1] == 0) dpi[1] = PIXEL_DPI;
                SKSizeI size = new SKSizeI();
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
        public static SKSizeI GetImageDimension(Stream is1, PictureType type)
        {
            SKSizeI size = new SKSizeI();

            switch (type)
            {
                case PictureType.JPEG:
                case PictureType.PNG:
                case PictureType.DIB:
                    //we can calculate the preferred size only for JPEG, PNG and BMP
                    //other formats like WMF, EMF and PICT are not supported in Java
                    byte[] data;
                    using (var ms = new MemoryStream())
                    {
                        is1.CopyTo(ms);
                        data = ms.ToArray();
                    }
                    using (SKBitmap img = SKBitmap.Decode(data))
                    {
                        if (img == null) return size;
                        int[] dpi = GetResolutionFromBytes(data);

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
         * Extract the DPI resolution from raw image bytes.
         * Supports JPEG (JFIF) and PNG (pHYs chunk).
         *
         * @return array of two elements: <code>{horizontalDpi, verticalDpi}</code>.
         * {0, 0} is returned when DPI cannot be determined.
         */
        public static int[] GetResolutionFromBytes(byte[] data)
        {
            if (data == null || data.Length < 4)
                return new int[] { 0, 0 };

            // JPEG: SOI marker FF D8
            if (data[0] == 0xFF && data[1] == 0xD8)
                return GetJpegDpi(data);

            // PNG: signature 89 50 4E 47 0D 0A 1A 0A
            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                return GetPngDpi(data);

            return new int[] { 0, 0 };
        }

        private static int[] GetJpegDpi(byte[] data)
        {
            int pos = 2; // skip SOI
            while (pos + 3 < data.Length)
            {
                if (data[pos] != 0xFF) break;
                byte marker = data[pos + 1];
                if (marker == 0xD9 || marker == 0xDA) break; // EOI or SOS

                int segLen = (data[pos + 2] << 8) | data[pos + 3];
                // APP0 JFIF
                if (marker == 0xE0 && segLen >= 16 && pos + segLen < data.Length)
                {
                    // Check "JFIF\0" identifier
                    if (data[pos + 4] == 'J' && data[pos + 5] == 'F' &&
                        data[pos + 6] == 'I' && data[pos + 7] == 'F' &&
                        data[pos + 8] == 0)
                    {
                        byte units = data[pos + 11];
                        int xDensity = (data[pos + 12] << 8) | data[pos + 13];
                        int yDensity = (data[pos + 14] << 8) | data[pos + 15];
                        if (units == 1) // pixels per inch
                            return new int[] { xDensity, yDensity };
                        else if (units == 2) // pixels per cm
                            return new int[] { (int)Math.Round(xDensity * 2.54), (int)Math.Round(yDensity * 2.54) };
                        // units == 0: aspect ratio only, no DPI
                    }
                }
                pos += 2 + segLen;
            }
            return new int[] { 0, 0 };
        }

        private static int[] GetPngDpi(byte[] data)
        {
            int pos = 8; // skip PNG signature
            while (pos + 12 <= data.Length)
            {
                int chunkLen = (data[pos] << 24) | (data[pos + 1] << 16) | (data[pos + 2] << 8) | data[pos + 3];
                if (pos + 8 + chunkLen > data.Length) break;

                // pHYs chunk: 9 bytes of data (4 + 4 + 1)
                if (data[pos + 4] == 'p' && data[pos + 5] == 'H' &&
                    data[pos + 6] == 'Y' && data[pos + 7] == 's' && chunkLen == 9)
                {
                    long xPpu = ((long)(data[pos + 8] & 0xFF) << 24) | ((long)(data[pos + 9] & 0xFF) << 16)
                              | ((long)(data[pos + 10] & 0xFF) << 8) | (long)(data[pos + 11] & 0xFF);
                    long yPpu = ((long)(data[pos + 12] & 0xFF) << 24) | ((long)(data[pos + 13] & 0xFF) << 16)
                              | ((long)(data[pos + 14] & 0xFF) << 8) | (long)(data[pos + 15] & 0xFF);
                    byte unit = data[pos + 16];
                    if (unit == 1) // metre
                        return new int[] { (int)Math.Round(xPpu * 0.0254), (int)Math.Round(yPpu * 0.0254) };
                }
                else if (data[pos + 4] == 'I' && data[pos + 5] == 'D' &&
                         data[pos + 6] == 'A' && data[pos + 7] == 'T')
                    break; // stop at IDAT

                pos += 4 + 4 + chunkLen + 4; // length + type + data + CRC
            }
            return new int[] { 0, 0 };
        }

        /**
         * @deprecated Use {@link #GetResolutionFromBytes(byte[])} instead.
         */
        public static int[] GetResolution(SKBitmap r)
        {
            // SKBitmap does not expose DPI metadata; return zeros so callers default to 96 DPI
            return new int[] { 0, 0 };
        }

        /**
         * Calculate and Set the preferred size (anchor) for this picture.
         *
         * @param scaleX the amount by which image width is multiplied relative to the original width.
         * @param scaleY the amount by which image height is multiplied relative to the original height.
         * @return the new Dimensions of the scaled picture in EMUs
         */
        public static SKSizeI SetPreferredSize(IPicture picture, double scaleX, double scaleY)
        {
            IClientAnchor anchor = picture.ClientAnchor;
            bool isHSSF = (anchor is HSSFClientAnchor);
            IPictureData data = picture.PictureData;
            ISheet sheet = picture.Sheet;

            // in pixel
            SKSizeI imgSize = GetImageDimension(new MemoryStream(data.Data), data.PictureType);
            // in emus
            SKSizeI anchorSize = ImageUtils.GetDimensionFromAnchor(picture);
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
                w *= 1d - anchor.Dx1 / 1024d;
            }
            else
            {
                w -= anchor.Dx1 / (double)Units.EMU_PER_PIXEL;
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
                h -= anchor.Dy1 / (double)Units.EMU_PER_PIXEL;
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

            SKSizeI dim = new SKSizeI(
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
        public static SKSizeI GetDimensionFromAnchor(IPicture picture)
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
                w -= anchor.Dx1 / (double)Units.EMU_PER_PIXEL;
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
                w += anchor.Dx2 / (double)Units.EMU_PER_PIXEL;
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
                h -= anchor.Dy1 / (double)Units.EMU_PER_PIXEL;
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
                h += anchor.Dy2 / (double)Units.EMU_PER_PIXEL;
            }

            w *= Units.EMU_PER_PIXEL;
            h *= Units.EMU_PER_PIXEL;

            return new SKSizeI((int)Math.Round(w), (int)Math.Round(h));
            //return new SKSizeI((int)w * Units.EMU_PER_PIXEL, (int)h * Units.EMU_PER_PIXEL);

        }


        public static double GetRowHeightInPixels(ISheet sheet, int rowNum)
        {
            IRow r = sheet.GetRow(rowNum);
            double points = (r == null) ? sheet.DefaultRowHeightInPoints : r.HeightInPoints;
            return Units.ToEMU(points) / Units.EMU_PER_PIXEL;
        }
    }

}
