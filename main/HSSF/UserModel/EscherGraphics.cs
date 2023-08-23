/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.HSSF.Util;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using SixLabors.ImageSharp;
    using SixLabors.ImageSharp.PixelFormats;
    using SixLabors.Fonts;

    /**
     * Translates Graphics calls into escher calls.  The translation Is lossy so
     * many features are not supported and some just aren't implemented yet.  If
     * in doubt test the specific calls you wish to make. Graphics calls are
     * always performed into an EscherGroup so one will need to be Created.
     * 
     * <b>Important:</b>
     * <blockquote>
     * One important concept worth considering Is that of font size.  One of the
     * difficulties in Converting Graphics calls into escher Drawing calls Is that
     * Excel does not have the concept of absolute pixel positions.  It measures
     * it's cell widths in 'Chars' and the cell heights in points.
     * Unfortunately it's not defined exactly what a type of Char it's
     * measuring.  Presumably this Is due to the fact that the Excel will be
     * using different fonts on different platforms or even within the same
     * platform.
     * 
     * Because of this constraint we've had to calculate the
     * verticalPointsPerPixel.  This the amount the font should be scaled by when
     * you Issue commands such as DrawString().  A good way to calculate this
     * Is to use the follow formula:
     * 
     * <pre>
     *      multipler = GroupHeightInPoints / heightOfGroup
     * </pre>
     * 
     * The height of the Group Is calculated fairly simply by calculating the
     * difference between the y coordinates of the bounding box of the shape.  The
     * height of the Group can be calculated by using a convenience called
     * <c>HSSFClientAnchor.GetAnchorHeightInPoints()</c>.
     * </blockquote>
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class EscherGraphics : IDisposable
    {
        private HSSFShapeGroup escherGroup;
        private HSSFWorkbook workbook;
        private float verticalPointsPerPixel = 1.0f;
        private float verticalPixelsPerPoint;
        private Rgb24 foreground;
        private Rgb24 background = new Rgb24(255, 255, 255);
        private Font font;
        private static POILogger Logger = POILogFactory.GetLogger(typeof(EscherGraphics));

        // Default dpi
        private static int dpi = 96;

        /**
         * Construct an escher graphics object.
         *
         * @param escherGroup           The escher Group to Write the graphics calls into.
         * @param workbook              The workbook we are using.
         * @param forecolor             The foreground color to use as default.
         * @param verticalPointsPerPixel    The font multiplier.  (See class description for information on how this works.).
         */
        public EscherGraphics(HSSFShapeGroup escherGroup, HSSFWorkbook workbook, Color forecolor, float verticalPointsPerPixel)
        {
            this.escherGroup = escherGroup;
            this.workbook = workbook;
            this.verticalPointsPerPixel = verticalPointsPerPixel;
            this.verticalPixelsPerPoint = 1 / verticalPointsPerPixel;
            this.font = new Font(SystemFonts.Get("Arial"), 10);
            this.foreground = forecolor;
            //        background = backcolor;
        }

        /**
         * Constructs an escher graphics object.
         *
         * @param escherGroup           The escher Group to Write the graphics calls into.
         * @param workbook              The workbook we are using.
         * @param foreground            The foreground color to use as default.
         * @param verticalPointsPerPixel    The font multiplier.  (See class description for information on how this works.).
         * @param font                  The font to use.
         */
        EscherGraphics(HSSFShapeGroup escherGroup, HSSFWorkbook workbook, Color foreground, Font font, float verticalPointsPerPixel)
        {
            this.escherGroup = escherGroup;
            this.workbook = workbook;
            this.foreground = foreground;
            //        this.background = background;
            this.font = font;
            this.verticalPointsPerPixel = verticalPointsPerPixel;
            this.verticalPixelsPerPoint = 1 / verticalPointsPerPixel;
        }

        //    /**
        //     * Constructs an escher graphics object.
        //     *
        //     * @param escherGroup           The escher Group to Write the graphics calls into.
        //     * @param workbook              The workbook we are using.
        //     * @param forecolor             The default foreground color.
        //     */
        //    public EscherGraphics( HSSFShapeGroup escherGroup, HSSFWorkbook workbook, Color forecolor)
        //    {
        //        this(escherGroup, workbook, forecolor, 1.0f);
        //    }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (null != font)
                {
                    font = null;
                }
            }
        }

        public void ClearRect(int x, int y, int width, int height)
        {
            Color color = foreground;
            SetColor(background);
            FillRect(x, y, width, height);
            SetColor(color);
        }

        public void ClipRect(int x, int y, int width, int height)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "clipRect not supported");
        }

        public void CopyArea(int x, int y, int width, int height, int dx, int dy)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "copyArea not supported");
        }

        public EscherGraphics Create()
        {
            EscherGraphics g = new EscherGraphics(escherGroup, workbook,
                    foreground, font, verticalPointsPerPixel);
            return g;
        }

        public void DrawArc(int x, int y, int width, int height,
                     int startAngle, int arcAngle)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "DrawArc not supported");
        }

        public bool DrawImage(Image img,
                          int dx1, int dy1, int dx2, int dy2,
                          int sx1, int sy1, int sx2, int sy2,
                          Color bgcolor)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "DrawImage not supported");

            throw new NotImplementedException();

            //return true;
        }

        public bool DrawImage(Image img,
                          int dx1, int dy1, int dx2, int dy2,
                          int sx1, int sy1, int sx2, int sy2)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "DrawImage not supported");

            throw new NotImplementedException();

            //return true;
        }

        public bool DrawImage(Image image, int i, int j, int k, int l, Color color)
        {
            return DrawImage(image, i, j, i + k, j + l, 0, 0, image.Width, image.Height, color);
        }

        public bool DrawImage(Image image, int i, int j, int k, int l)
        {
            return DrawImage(image, i, j, i + k, j + l, 0, 0, image.Width, image.Height);
        }

        public bool DrawImage(Image image, int i, int j, Color color)
        {
            return DrawImage(image, i, j, image.Width, image.Height, color);
        }

        public bool DrawImage(Image image, int i, int j)
        {
            return DrawImage(image, i, j, image.Width, image.Height);
        }

        public void DrawLine(int x1, int y1, int x2, int y2)
        {
            DrawLine(x1, y1, x2, y2, 0);
        }

        public void DrawLine(int x1, int y1, int x2, int y2, int width)
        {
            HSSFSimpleShape shape = escherGroup.CreateShape(new HSSFChildAnchor(x1, y1, x2, y2));
            shape.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            shape.LineWidth = (width);
            shape.SetLineStyleColor(foreground.R, foreground.G, foreground.B);
        }

        public void DrawOval(int x, int y, int width, int height)
        {
            HSSFSimpleShape shape = escherGroup.CreateShape(new HSSFChildAnchor(x, y, x + width, y + height));
            shape.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_OVAL);
            shape.LineWidth = 0;
            shape.SetLineStyleColor(foreground.R, foreground.G, foreground.B);
            shape.IsNoFill = (true);
        }

        public void DrawPolygon(int[] xPoints, int[] yPoints, int nPoints)
        {
            int right = FindBiggest(xPoints);
            int bottom = FindBiggest(yPoints);
            int left = FindSmallest(xPoints);
            int top = FindSmallest(yPoints);
            HSSFPolygon shape = escherGroup.CreatePolygon(new HSSFChildAnchor(left, top, right, bottom));
            shape.SetPolygonDrawArea(right - left, bottom - top);
            shape.SetPoints(AddToAll(xPoints, -left), AddToAll(yPoints, -top));
            shape.SetLineStyleColor(foreground.R, foreground.G, foreground.B);
            shape.LineWidth = (0);
            shape.IsNoFill = (true);
        }

        private int[] AddToAll(int[] values, int amount)
        {
            int[] result = new int[values.Length];
            for (int i = 0; i < values.Length; i++)
                result[i] = values[i] + amount;
            return result;
        }

        public void DrawPolyline(int[] xPoints, int[] yPoints,
                          int nPoints)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "DrawPolyline not supported");
        }

        public void DrawRect(int x, int y, int width, int height)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "DrawRect not supported");
        }

        public void DrawRoundRect(int x, int y, int width, int height,
                           int arcWidth, int arcHeight)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "DrawRoundRect not supported");
        }

        public void DrawString(String str, int x, int y)
        {
            if (string.IsNullOrEmpty(str))
                return;
            // TODO-Fonts: Fallback for missing font
            Font excelFont = new Font(SystemFonts.Get(font.Name.Equals("SansSerif") ? "Arial" : font.Name),
                (int)(font.Size / verticalPixelsPerPoint), font.FontMetrics.Description.Style);
            {
                var textOptions = new TextOptions(excelFont) { Dpi = dpi };
                int width = (int)((TextMeasurer.MeasureSize(str, textOptions).Width * 8) + 12);
                int height = (int)((font.Size / verticalPixelsPerPoint) + 6) * 2;
                y -= Convert.ToInt32((font.Size / verticalPixelsPerPoint) + 2 * verticalPixelsPerPoint);    // we want to Draw the shape from the top-left
                HSSFTextbox textbox = escherGroup.CreateTextbox(new HSSFChildAnchor(x, y, x + width, y + height));
                textbox.IsNoFill = (true);
                textbox.LineStyle = LineStyle.None;
                HSSFRichTextString s = new HSSFRichTextString(str);
                HSSFFont hssfFont = MatchFont(excelFont);
                s.ApplyFont(hssfFont);
                textbox.String = (s);
            }
        }

        private HSSFFont MatchFont(Font font)
        {
            HSSFColor hssfColor = workbook.GetCustomPalette()
                    .FindColor((byte)foreground.R, (byte)foreground.G, (byte)foreground.B);
            if (hssfColor == null)
                hssfColor = workbook.GetCustomPalette().FindSimilarColor((byte)foreground.R, (byte)foreground.G, (byte)foreground.B);
            bool bold = font.IsBold;
            bool italic = font.IsItalic;
            HSSFFont hssfFont = (HSSFFont)workbook.FindFont(bold ? (short)NPOI.SS.UserModel.FontBoldWeight.Bold : (short)NPOI.SS.UserModel.FontBoldWeight.Normal,
                        hssfColor.Indexed,
                        (short)(font.Size * 20),
                        font.Name,
                        italic,
                        false,
                        (short)NPOI.SS.UserModel.FontSuperScript.None,
                        (byte)NPOI.SS.UserModel.FontUnderlineType.None
                        );
            if (hssfFont == null)
            {
                hssfFont = (HSSFFont)workbook.CreateFont();
                hssfFont.Boldweight = (short)(bold ? NPOI.SS.UserModel.FontBoldWeight.Bold : 0);
                hssfFont.Color = (hssfColor.Indexed);
                hssfFont.FontHeight = ((short)(font.Size * 20));
                hssfFont.FontName = font.Name;
                hssfFont.IsItalic = (italic);
                hssfFont.IsStrikeout = (false);
                hssfFont.TypeOffset = 0;
                hssfFont.Underline = 0;
            }

            return hssfFont;
        }


        //public void DrawString(AttributedCharIEnumerator iterator,
        //                                int x, int y)
        //{
        //    if (Logger.Check(POILogger.WARN))
        //        Logger.Log(POILogger.WARN, "DrawString not supported");
        //}

        public void FillArc(int x, int y, int width, int height,
                     int startAngle, int arcAngle)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "FillArc not supported");
        }

        public void FillOval(int x, int y, int width, int height)
        {
            HSSFSimpleShape shape = escherGroup.CreateShape(new HSSFChildAnchor(x, y, x + width, y + height));
            shape.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_OVAL);
            shape.LineStyle = LineStyle.None;
            shape.SetFillColor(foreground.R, foreground.G, foreground.B);
            shape.SetLineStyleColor(foreground.R, foreground.G, foreground.B);
            shape.IsNoFill = (false);
        }

        /**
         * Fills a (closed) polygon, as defined by a pair of arrays, which
         *  hold the <i>x</i> and <i>y</i> coordinates.
         * 
         * This Draws the polygon, with <c>nPoint</c> line segments.
         * The first <c>nPoint - 1</c> line segments are
         *  Drawn between sequential points 
         *  (<c>xPoints[i],yPoints[i],xPoints[i+1],yPoints[i+1]</c>).
         * The line segment Is a closing one, from the last point to 
         *  the first (assuming they are different).
         * 
         * The area inside of the polygon Is defined by using an
         *  even-odd Fill rule (also known as the alternating rule), and 
         *  the area inside of it Is Filled.
         * @param xPoints array of the <c>x</c> coordinates.
         * @param yPoints array of the <c>y</c> coordinates.
         * @param nPoints the total number of points in the polygon.
         * @see   java.awt.Graphics#DrawPolygon(int[], int[], int)
         */
        public void FillPolygon(int[] xPoints, int[] yPoints,
                         int nPoints)
        {
            int right = FindBiggest(xPoints);
            int bottom = FindBiggest(yPoints);
            int left = FindSmallest(xPoints);
            int top = FindSmallest(yPoints);
            HSSFPolygon shape = escherGroup.CreatePolygon(new HSSFChildAnchor(left, top, right, bottom));
            shape.SetPolygonDrawArea(right - left, bottom - top);
            shape.SetPoints(AddToAll(xPoints, -left), AddToAll(yPoints, -top));
            shape.SetLineStyleColor(foreground.R, foreground.G, foreground.B);
            shape.SetFillColor(foreground.R, foreground.G, foreground.B);
        }

        private int FindBiggest(int[] values)
        {
            int result = Int32.MinValue;
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] > result)
                    result = values[i];
            }
            return result;
        }

        private int FindSmallest(int[] values)
        {
            int result = Int32.MaxValue;
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] < result)
                    result = values[i];
            }
            return result;
        }

        public void FillRect(int x, int y, int width, int height)
        {
            HSSFSimpleShape shape = escherGroup.CreateShape(new HSSFChildAnchor(x, y, x + width, y + height));
            shape.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
            shape.LineStyle = LineStyle.None;
            shape.SetFillColor(foreground.R, foreground.G, foreground.B);
            shape.SetLineStyleColor(foreground.R, foreground.G, foreground.B);
        }

        public void FillRoundRect(int x, int y, int width, int height,
                           int arcWidth, int arcHeight)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "FillRoundRect not supported");
        }

        public Rectangle Clip
        {
            get
            {
                return ClipBounds;
            }
        }

        public Rectangle ClipBounds
        {
            get
            {
                return Rectangle.Empty;
            }
        }

        public Color Color
        {
            get
            {
                return foreground;
            }
        }

        public Font Font
        {
            get
            {
                return font;
            }
        }

        //public FontMetrics GetFontMetrics(Font f)
        //{
        //    return GetFontMetrics(f);
        //}

        public void SetClip(int x, int y, int width, int height)
        {
            SetClip(((new Rectangle(x, y, width, height))));
        }

        public void SetClip(Rectangle shape)
        {
            // ignore... not implemented
            throw new NotImplementedException();
        }

        public void SetColor(Color color)
        {
            foreground = color;
        }

        public void SetFont(Font f)
        {
            font = f;
        }

        public void SetPaintMode()
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "SetPaintMode not supported");

            throw new NotImplementedException();
        }

        public void SetXORMode(Color color)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "SetXORMode not supported");

            throw new NotImplementedException();
        }

        public void Translate(int x, int y)
        {
            if (Logger.Check(POILogger.WARN))
                Logger.Log(POILogger.WARN, "translate not supported");

            throw new NotImplementedException();
        }

        public Color Background
        {
            get
            {
                return background;
            }
            set 
            {
                this.background = value;
            }
        }

        HSSFShapeGroup GetEscherGraphics()
        {
                return escherGroup;

        }
    }

}