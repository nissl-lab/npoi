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

namespace NPOI.XSSF.Streaming
{
    using System.IO;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.OpenXmlFormats.Dml.Spreadsheet;
    using SixLabors.ImageSharp;

    /// <summary>
    /// <para>
    /// Streaming version of Picture.
    /// Most of the code is a copy of the non-streaming XSSFPicture code.
    /// This is necessary as a private method GetRowHeightInPixels of that class needs to be Changed, which is called by a method call chain nested several levels.
    /// </para>
    /// <para>
    /// The main change is to access the rows in the SXSSF sheet, not the always empty rows in the XSSF sheet when Checking the row heights.
    /// </para>
    /// </summary>
    public class SXSSFPicture : IPicture
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(SXSSFPicture));
        /// <summary>
        /// <para>
        /// Column width measured as the number of characters of the maximum digit width of the
        /// numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. There are 4 pixels of margin
        /// pAdding (two on each side), plus 1 pixel pAdding for the gridlines.
        /// </para>
        /// <para>
        /// This value is the same for default font in Office 2007 (Calibry) and Office 2003 and earlier (Arial)
        /// </para>
        /// </summary>
        private static float DEFAULT_COLUMN_WIDTH = 9.140625f;

        private SXSSFWorkbook _wb;
        private XSSFPicture _picture;

        public SXSSFPicture(SXSSFWorkbook _wb, XSSFPicture _picture)
        {
            this._wb = _wb;
            this._picture = _picture;
        }

        /// <summary>
        /// <para>
        /// Return the underlying CTPicture bean that holds all properties for this picture
        /// </para>
        /// <para>
        /// </para>
        /// </summary>
        /// <return>underlying CTPicture bean/// </return>

        public CT_Picture GetCTPicture()
        {
            return _picture.GetCTPicture();
        }

        /// <summary>
        /// <para>
        /// Reset the image to the original size.
        /// </para>
        /// <para>
        /// <p>
        /// Please note, that this method works correctly only for workbooks
        /// with the default font size (Calibri 11pt for .xlsx).
        /// If the default font is Changed the resized image can be streched vertically or horizontally.
        /// </p>
        /// </para>
        /// </summary>
        public void Resize()
        {
            Resize(1.0);
        }

        /// <summary>
        /// <para>
        /// Reset the image to the original size.
        /// <p>
        /// Please note, that this method works correctly only for workbooks
        /// with the default font size (Calibri 11pt for .xlsx).
        /// If the default font is Changed the resized image can be streched vertically or horizontally.
        /// </p>
        /// </para>
        /// <para>
        /// </para>
        /// </summary>
        /// <param name="scale">the amount by which image dimensions are multiplied relative to the original size.
        ///<code>resize(1.0)</code> Sets the original size, <code>resize(0.5)</code> resize to 50% of the original,
        ///<code>resize(2.0)</code> resizes to 200% of the original.
        /// </param>
        public void Resize(double scale)
        {
            XSSFClientAnchor anchor = (XSSFClientAnchor)ClientAnchor;

            XSSFClientAnchor pref = GetPreferredSize(scale);

            int row2 = anchor.Row1 + (pref.Row2 - pref.Row1);
            int col2 = anchor.Col1 + (pref.Col2 - pref.Col1);

            anchor.Col2 = (/*setter*/col2);
            anchor.Dx1 = (/*setter*/0);
            anchor.Dx2 = (/*setter*/pref.Dx2);

            anchor.Row2 = (/*setter*/row2);
            anchor.Dy1 = (/*setter*/0);
            anchor.Dy2 = (/*setter*/pref.Dy2);
        }

        /// <summary>
        /// <para>
        /// Calculate the preferred size for this picture.
        /// </para>
        /// <para>
        /// </para>
        /// </summary>
        /// <return>with the preferred size for this image/// </return>
        public IClientAnchor GetPreferredSize()
        {
            return GetPreferredSize(1.0);
        }

        /// <summary>
        /// <para>
        /// Calculate the preferred size for this picture.
        /// </para>
        /// <para>
        /// </para>
        /// </summary>
        /// <param name="scale">the amount by which image dimensions are multiplied relative to the original size./// </param>
        /// <return>with the preferred size for this image/// </return>
        public XSSFClientAnchor GetPreferredSize(double scale)
        {
            XSSFClientAnchor anchor = (XSSFClientAnchor)ClientAnchor;

            XSSFPictureData data = (XSSFPictureData)PictureData;
            Size size = GetImageDimension(data.GetPackagePart(), data.PictureType);
            double scaledWidth = size.Width * scale;
            double scaledHeight = size.Height * scale;

            float w = 0;
            int col2 = anchor.Col1 - 1;

            while (w <= scaledWidth)
            {
                w += GetColumnWidthInPixels(++col2);
            }

            //assert(w > scaledWidth);
            double cw = GetColumnWidthInPixels(col2);
            double deltaW = w - scaledWidth;
            int dx2 = (int)(XSSFShape.EMU_PER_PIXEL * (cw - deltaW));

            anchor.Col2 = (/*setter*/col2);
            anchor.Dx2 = (/*setter*/dx2);

            double h = 0;
            int row2 = anchor.Row1 - 1;

            while (h <= scaledHeight)
            {
                h += GetRowHeightInPixels(++row2);
            }

            //assert(h > scaledHeight);
            double ch = GetRowHeightInPixels(row2);
            double deltaH = h - scaledHeight;
            int dy2 = (int)(XSSFShape.EMU_PER_PIXEL * (ch - deltaH));
            anchor.Row2 = (/*setter*/row2);
            anchor.Dy2 = (/*setter*/dy2);

            CT_PositiveSize2D size2d = GetCTPicture().spPr.xfrm.ext;
            size2d.cx = (/*setter*/(long)(scaledWidth * XSSFShape.EMU_PER_PIXEL));
            size2d.cy = (/*setter*/(long)(scaledHeight * XSSFShape.EMU_PER_PIXEL));

            return anchor;
        }

        private float GetColumnWidthInPixels(int columnIndex)
        {
            XSSFSheet sheet = (XSSFSheet)Sheet;

            CT_Col col = sheet.GetColumnHelper().GetColumn(columnIndex, false);
            double numChars = col == null || !col.IsSetWidth() ? DEFAULT_COLUMN_WIDTH : col.width;

            return (float)numChars * XSSFWorkbook.DEFAULT_CHARACTER_WIDTH;
        }

        private float GetRowHeightInPixels(int rowIndex)
        {
            // THE FOLLOWING THREE LINES ARE THE MAIN CHANGE Compared to the non-streaming version: use the SXSSF sheet,
            // not the XSSF sheet (which never contais rows when using SXSSF)
            XSSFSheet xssfSheet = (XSSFSheet)Sheet;
            SXSSFSheet sheet = _wb.GetSXSSFSheet(xssfSheet);
            IRow row = sheet.GetRow(rowIndex);
            float height = row != null ? row.HeightInPoints : sheet.DefaultRowHeightInPoints;
            return height * XSSFShape.PIXEL_DPI / XSSFShape.POINT_DPI;
        }
        /// <summary>
        /// <para>
        /// Return the dimension of this image
        /// </para>
        /// <para>
        /// </para>
        /// </summary>
        /// <param name="part">the package part holding raw picture data/// </param>
        /// <param name="type">type of the picture: {@link Workbook#PICTURE_TYPE_JPEG},
        ///{@link Workbook#PICTURE_TYPE_PNG} or {@link Workbook#PICTURE_TYPE_DIB}
        /// </param>
        /// 
        /// <return>dimension in pixels/// </return>
        protected static Size GetImageDimension(PackagePart part, PictureType type)
        {
            try
            {
                return ImageUtils.GetImageDimension(part.GetInputStream(), type);
            }
            catch (IOException e)
            {
                //return a "singulariry" if ImageIO failed to read the image
                logger.Log(POILogger.WARN, e);
                return new Size();
            }
        }

        /// <summary>
        /// <para>
        /// Return picture data for this shape
        /// </para>
        /// <para>
        /// </para>
        /// </summary>
        /// <return>data for this shape/// </return>
        public IPictureData PictureData
        {
            get
            {
                return _picture.PictureData;
            }
        }

        protected OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties GetShapeProperties()
        {
            return GetCTPicture().spPr;
        }
        public XSSFAnchor GetAnchor()
        {
            return _picture.GetAnchor();
        }
        public void Resize(double scaleX, double scaleY)
        {
            _picture.Resize(scaleX, scaleY);
        }
        public IClientAnchor GetPreferredSize(double scaleX, double scaleY)
        {
            return _picture.GetPreferredSize(scaleX, scaleY);
        }
        public Size GetImageDimension()
        {
            return _picture.GetImageDimension();
        }
        public IClientAnchor ClientAnchor
        {
            get
            {
                XSSFAnchor a = GetAnchor();
                return (a is XSSFClientAnchor) ? (XSSFClientAnchor)a : null;
            }
        }

        public XSSFDrawing GetDrawing()
        {
            return _picture.GetDrawing();
        }
        public ISheet Sheet
        {
            get
            {
                return _picture.Sheet;
            }
        }
        public string GetShapeName()
        {
            return _picture.Name;
        }
        public IShape GetParent()
        {
            return _picture.Parent;
        }
        public bool IsNoFill
        {
            get { return _picture.IsNoFill; }
            set { _picture.IsNoFill = value;}
        }

        public void SetFillColor(int red, int green, int blue)
        {
            _picture.SetFillColor(red, green, blue);
        }
        public void SetLineStyleColor(int red, int green, int blue)
        {
            _picture.SetLineStyleColor(red, green, blue);
        }
    }
}

