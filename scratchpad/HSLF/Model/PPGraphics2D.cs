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












using NPOI.HSLF.usermodel.RichTextRun;
using NPOI.HSLF.exceptions.HSLFException;
using NPOI.util.POILogger;
using NPOI.util.POILogFactory;

/**
 * Translates Graphics2D calls into PowerPoint.
 *
 * @author Yegor Kozlov
 */
public class PPGraphics2D : Graphics2D : Cloneable {

    protected POILogger log = POILogFactory.GetLogger(this.GetClass());

    //The ppt object to write into.
    private ShapeGroup _group;

    private AffineTransform _transform;
    private Stroke _stroke;
    private Paint _paint;
    private Font _font;
    private Color _foreground;
    private Color _background;
    private RenderingHints _hints;

    /**
     * Construct Java Graphics object which translates graphic calls in ppt Drawing layer.
     *
     * @param group           The shape group to write the graphics calls into.
     */
    public PPGraphics2D(ShapeGroup group){
        this._group = group;

        _transform = new AffineTransform();
        _stroke = new BasicStroke();
        _paint = Color.black;
        _font = new Font("Arial", Font.PLAIN, 12);
        _background = Color.black;
        _foreground = Color.white;
        _hints = new RenderingHints(null);
    }

    /**
     * @return  the shape group being used for Drawing
     */
    public ShapeGroup GetShapeGroup(){
        return _group;
    }

    /**
     * Gets the current font.
     * @return    this graphics context's current font.
     * @see       java.awt.Font
     * @see       java.awt.Graphics#setFont(Font)
     */
    public Font GetFont(){
        return _font;
    }

    /**
     * Sets this graphics context's font to the specified font.
     * All subsequent text operations using this graphics context
     * use this font.
     * @param  font   the font.
     * @see     java.awt.Graphics#getFont
     * @see     java.awt.Graphics#drawString(java.lang.String, int, int)
     * @see     java.awt.Graphics#drawBytes(byte[], int, int, int, int)
     * @see     java.awt.Graphics#drawChars(char[], int, int, int, int)
    */
    public void SetFont(Font font){
        this._font = font;
    }

    /**
     * Gets this graphics context's current color.
     * @return    this graphics context's current color.
     * @see       java.awt.Color
     * @see       java.awt.Graphics#setColor
     */
     public Color GetColor(){
        return _foreground;
    }

    /**
     * Sets this graphics context's current color to the specified
     * color. All subsequent graphics operations using this graphics
     * context use this specified color.
     * @param     c   the new rendering color.
     * @see       java.awt.Color
     * @see       java.awt.Graphics#getColor
     */
    public void SetColor(Color c) {
        SetPaint(c);
    }

    /**
     * Returns the current <code>Stroke</code> in the
     * <code>Graphics2D</code> context.
     * @return the current <code>Graphics2D</code> <code>Stroke</code>,
     *                 which defines the line style.
     * @see #setStroke
     */
    public Stroke GetStroke(){
        return _stroke;
    }

    /**
     * Sets the <code>Stroke</code> for the <code>Graphics2D</code> context.
     * @param s the <code>Stroke</code> object to be used to stroke a
     * <code>Shape</code> during the rendering process
     */
    public void SetStroke(Stroke s){
        this._stroke = s;
    }

    /**
     * Returns the current <code>Paint</code> of the
     * <code>Graphics2D</code> context.
     * @return the current <code>Graphics2D</code> <code>Paint</code>,
     * which defines a color or pattern.
     * @see #setPaint
     * @see java.awt.Graphics#setColor
     */
    public Paint GetPaint(){
        return _paint;
    }

    /**
     * Sets the <code>Paint</code> attribute for the
     * <code>Graphics2D</code> context.  Calling this method
     * with a <code>null</code> <code>Paint</code> object does
     * not have any effect on the current <code>Paint</code> attribute
     * of this <code>Graphics2D</code>.
     * @param paint the <code>Paint</code> object to be used to generate
     * color during the rendering Process, or <code>null</code>
     * @see java.awt.Graphics#setColor
     */
     public void SetPaint(Paint paint){
        if(paint == null) return;

        this._paint = paint;
        if (paint is Color) _foreground = (Color)paint;
    }

    /**
     * Returns a copy of the current <code>Transform</code> in the
     * <code>Graphics2D</code> context.
     * @return the current <code>AffineTransform</code> in the
     *             <code>Graphics2D</code> context.
     * @see #_transform
     * @see #setTransform
     */
    public AffineTransform GetTransform(){
        return new AffineTransform(_transform);
    }

    /**
     * Sets the <code>Transform</code> in the <code>Graphics2D</code>
     * context.
     * @param Tx the <code>AffineTransform</code> object to be used in the
     * rendering process
     * @see #_transform
     * @see AffineTransform
     */
    public void SetTransform(AffineTransform Tx) {
        _transform = new AffineTransform(Tx);
    }

    /**
     * Strokes the outline of a <code>Shape</code> using the Settings of the
     * current <code>Graphics2D</code> context.  The rendering attributes
     * applied include the <code>Clip</code>, <code>Transform</code>,
     * <code>Paint</code>, <code>Composite</code> and
     * <code>Stroke</code> attributes.
     * @param shape the <code>Shape</code> to be rendered
     * @see #setStroke
     * @see #setPaint
     * @see java.awt.Graphics#setColor
     * @see #_transform
     * @see #setTransform
     * @see #clip
     * @see #setClip
     * @see #setComposite
     */
    public void Draw(Shape shape){
        GeneralPath path = new GeneralPath(_transform.CreateTransformedShape(shape));
        Freeform p = new Freeform(_group);
        p.SetPath(path);
        p.GetFill().SetForegroundColor(null);
        ApplyStroke(p);
        _group.AddShape(p);
    }

    /**
     * Renders the text specified by the specified <code>String</code>,
     * using the current text attribute state in the <code>Graphics2D</code> context.
     * The baseline of the first character is at position
     * (<i>x</i>,&nbsp;<i>y</i>) in the User Space.
     * The rendering attributes applied include the <code>Clip</code>,
     * <code>Transform</code>, <code>Paint</code>, <code>Font</code> and
     * <code>Composite</code> attributes. For characters in script systems
     * such as Hebrew and Arabic, the glyphs can be rendered from right to
     * left, in which case the coordinate supplied is the location of the
     * leftmost character on the baseline.
     * @param s the <code>String</code> to be rendered
     * @param x the x coordinate of the location where the
     * <code>String</code> should be rendered
     * @param y the y coordinate of the location where the
     * <code>String</code> should be rendered
     * @throws NullPointerException if <code>str</code> is
     *         <code>null</code>
     * @see #setPaint
     * @see java.awt.Graphics#setColor
     * @see java.awt.Graphics#setFont
     * @see #setTransform
     * @see #setComposite
     * @see #setClip
     */
    public void DrawString(String s, float x, float y) {
        TextBox txt = new TextBox(_group);
        txt.GetTextRun().supplySlideShow(_group.Sheet.GetSlideShow());
        txt.GetTextRun().SetSheet(_group.Sheet);
        txt.SetText(s);

        RichTextRun rt = txt.GetTextRun().GetRichTextRuns()[0];
        rt.SetFontSize(_font.GetSize());
        rt.SetFontName(_font.GetFamily());

        if (getColor() != null) rt.SetFontColor(getColor());
        if (_font.IsBold()) rt.SetBold(true);
        if (_font.IsItalic()) rt.SetItalic(true);

        txt.SetMarginBottom(0);
        txt.SetMarginTop(0);
        txt.SetMarginLeft(0);
        txt.SetMarginRight(0);
        txt.SetWordWrap(TextBox.WrapNone);
        txt.SetHorizontalAlignment(TextBox.AlignLeft);
        txt.SetVerticalAlignment(TextBox.AnchorMiddle);


        TextLayout layout = new TextLayout(s, _font, GetFontRenderContext());
        float ascent = layout.GetAscent();

        float width = (float) Math.floor(layout.GetAdvance());
        /**
         * Even if top and bottom margins are Set to 0 PowerPoint
         * always Sets extra space between the text and its bounding box.
         *
         * The approximation height = ascent*2 works good enough in most cases
         */
        float height = ascent * 2;

        /*
          In powerpoint anchor of a shape is its top left corner.
          Java graphics Sets string coordinates by the baseline of the first character
          so we need to shift up by the height of the textbox
        */
        y -= height / 2 + ascent / 2;

        /*
          In powerpoint anchor of a shape is its top left corner.
          Java graphics Sets string coordinates by the baseline of the first character
          so we need to shift down by the height of the textbox
        */
        txt.SetAnchor(new Rectangle2D.Float(x, y, width, height));

        _group.AddShape(txt);
    }

    /**
     * Fills the interior of a <code>Shape</code> using the Settings of the
     * <code>Graphics2D</code> context. The rendering attributes applied
     * include the <code>Clip</code>, <code>Transform</code>,
     * <code>Paint</code>, and <code>Composite</code>.
     * @param shape the <code>Shape</code> to be Filled
     * @see #setPaint
     * @see java.awt.Graphics#setColor
     * @see #_transform
     * @see #setTransform
     * @see #setComposite
     * @see #clip
     * @see #setClip
     */
    public void Fill(Shape shape){
        GeneralPath path = new GeneralPath(_transform.CreateTransformedShape(shape));
        Freeform p = new Freeform(_group);
        p.SetPath(path);
        ApplyPaint(p);
        p.SetLineColor(null);   //Fills must be "No Line"
        _group.AddShape(p);
    }

    /**
     * Translates the origin of the graphics context to the point
     * (<i>x</i>,&nbsp;<i>y</i>) in the current coordinate system.
     * Modifies this graphics context so that its new origin corresponds
     * to the point (<i>x</i>,&nbsp;<i>y</i>) in this graphics context's
     * original coordinate system.  All coordinates used in subsequent
     * rendering operations on this graphics context will be relative
     * to this new origin.
     * @param  x   the <i>x</i> coordinate.
     * @param  y   the <i>y</i> coordinate.
     */
    public void translate(int x, int y){
        _transform.translate(x, y);
    }

    /**
     * Intersects the current <code>Clip</code> with the interior of the
     * specified <code>Shape</code> and Sets the <code>Clip</code> to the
     * resulting intersection.  The specified <code>Shape</code> is
     * transformed with the current <code>Graphics2D</code>
     * <code>Transform</code> before being intersected with the current
     * <code>Clip</code>.  This method is used to make the current
     * <code>Clip</code> smaller.
     * To make the <code>Clip</code> larger, use <code>setClip</code>.
     * The <i>user clip</i> modified by this method is independent of the
     * clipping associated with device bounds and visibility.  If no clip has
     * previously been Set, or if the clip has been Cleared using
     * {@link java.awt.Graphics#setClip(Shape) SetClip} with a
     * <code>null</code> argument, the specified <code>Shape</code> becomes
     * the new user clip.
     * @param s the <code>Shape</code> to be intersected with the current
     *          <code>Clip</code>.  If <code>s</code> is <code>null</code>,
     *          this method Clears the current <code>Clip</code>.
     */
    public void clip(Shape s){
        log.log(POILogger.WARN, "Not implemented");
    }

    /**
     * Gets the current clipping area.
     * This method returns the user clip, which is independent of the
     * clipping associated with device bounds and window visibility.
     * If no clip has previously been Set, or if the clip has been
     * Cleared using <code>setClip(null)</code>, this method returns
     * <code>null</code>.
     * @return      a <code>Shape</code> object representing the
     *              current clipping area, or <code>null</code> if
     *              no clip is Set.
     * @see         java.awt.Graphics#getClipBounds()
     * @see         java.awt.Graphics#clipRect
     * @see         java.awt.Graphics#setClip(int, int, int, int)
     * @see         java.awt.Graphics#setClip(Shape)
     * @since       JDK1.1
     */
    public Shape GetClip(){
        log.log(POILogger.WARN, "Not implemented");
        return null;
    }

    /**
     * Concatenates the current <code>Graphics2D</code>
     * <code>Transform</code> with a scaling transformation
     * Subsequent rendering is resized according to the specified scaling
     * factors relative to the previous scaling.
     * This is equivalent to calling <code>transform(S)</code>, where S is an
     * <code>AffineTransform</code> represented by the following matrix:
     * <pre>
     *          [   sx   0    0   ]
     *          [   0    sy   0   ]
     *          [   0    0    1   ]
     * </pre>
     * @param sx the amount by which X coordinates in subsequent
     * rendering operations are multiplied relative to previous
     * rendering operations.
     * @param sy the amount by which Y coordinates in subsequent
     * rendering operations are multiplied relative to previous
     * rendering operations.
     */
    public void scale(double sx, double sy){
        _transform.scale(sx, sy);
    }

    /**
     * Draws an outlined round-cornered rectangle using this graphics
     * context's current color. The left and right edges of the rectangle
     * are at <code>x</code> and <code>x&nbsp;+&nbsp;width</code>,
     * respectively. The top and bottom edges of the rectangle are at
     * <code>y</code> and <code>y&nbsp;+&nbsp;height</code>.
     * @param      x the <i>x</i> coordinate of the rectangle to be Drawn.
     * @param      y the <i>y</i> coordinate of the rectangle to be Drawn.
     * @param      width the width of the rectangle to be Drawn.
     * @param      height the height of the rectangle to be Drawn.
     * @param      arcWidth the horizontal diameter of the arc
     *                    at the four corners.
     * @param      arcHeight the vertical diameter of the arc
     *                    at the four corners.
     * @see        java.awt.Graphics#FillRoundRect
     */
    public void DrawRoundRect(int x, int y, int width, int height,
                              int arcWidth, int arcHeight){
        RoundRectangle2D rect = new RoundRectangle2D.Float(x, y, width, height, arcWidth, arcHeight);
        Draw(rect);
     }

    /**
     * Draws the text given by the specified string, using this
     * graphics context's current font and color. The baseline of the
     * first character is at position (<i>x</i>,&nbsp;<i>y</i>) in this
     * graphics context's coordinate system.
     * @param       str      the string to be Drawn.
     * @param       x        the <i>x</i> coordinate.
     * @param       y        the <i>y</i> coordinate.
     * @see         java.awt.Graphics#drawBytes
     * @see         java.awt.Graphics#drawChars
     */
    public void DrawString(String str, int x, int y){
        DrawString(str, (float)x, (float)y);
    }

    /**
     * Fills an oval bounded by the specified rectangle with the
     * current color.
     * @param       x the <i>x</i> coordinate of the upper left corner
     *                     of the oval to be Filled.
     * @param       y the <i>y</i> coordinate of the upper left corner
     *                     of the oval to be Filled.
     * @param       width the width of the oval to be Filled.
     * @param       height the height of the oval to be Filled.
     * @see         java.awt.Graphics#drawOval
     */
    public void FillOval(int x, int y, int width, int height){
        Ellipse2D oval = new Ellipse2D.Float(x, y, width, height);
        Fill(oval);
    }

    /**
     * Fills the specified rounded corner rectangle with the current color.
     * The left and right edges of the rectangle
     * are at <code>x</code> and <code>x&nbsp;+&nbsp;width&nbsp;-&nbsp;1</code>,
     * respectively. The top and bottom edges of the rectangle are at
     * <code>y</code> and <code>y&nbsp;+&nbsp;height&nbsp;-&nbsp;1</code>.
     * @param       x the <i>x</i> coordinate of the rectangle to be Filled.
     * @param       y the <i>y</i> coordinate of the rectangle to be Filled.
     * @param       width the width of the rectangle to be Filled.
     * @param       height the height of the rectangle to be Filled.
     * @param       arcWidth the horizontal diameter
     *                     of the arc at the four corners.
     * @param       arcHeight the vertical diameter
     *                     of the arc at the four corners.
     * @see         java.awt.Graphics#drawRoundRect
     */
    public void FillRoundRect(int x, int y, int width, int height,
                              int arcWidth, int arcHeight){

        RoundRectangle2D rect = new RoundRectangle2D.Float(x, y, width, height, arcWidth, arcHeight);
        Fill(rect);
    }

    /**
     * Fills a circular or elliptical arc covering the specified rectangle.
     * <p>
     * The resulting arc begins at <code>startAngle</code> and extends
     * for <code>arcAngle</code> degrees.
     * Angles are interpreted such that 0&nbsp;degrees
     * is at the 3&nbsp;o'clock position.
     * A positive value indicates a counter-clockwise rotation
     * while a negative value indicates a clockwise rotation.
     * <p>
     * The center of the arc is the center of the rectangle whose origin
     * is (<i>x</i>,&nbsp;<i>y</i>) and whose size is specified by the
     * <code>width</code> and <code>height</code> arguments.
     * <p>
     * The resulting arc covers an area
     * <code>width&nbsp;+&nbsp;1</code> pixels wide
     * by <code>height&nbsp;+&nbsp;1</code> pixels tall.
     * <p>
     * The angles are specified relative to the non-square extents of
     * the bounding rectangle such that 45 degrees always falls on the
     * line from the center of the ellipse to the upper right corner of
     * the bounding rectangle. As a result, if the bounding rectangle is
     * noticeably longer in one axis than the other, the angles to the
     * start and end of the arc segment will be skewed farther along the
     * longer axis of the bounds.
     * @param        x the <i>x</i> coordinate of the
     *                    upper-left corner of the arc to be Filled.
     * @param        y the <i>y</i>  coordinate of the
     *                    upper-left corner of the arc to be Filled.
     * @param        width the width of the arc to be Filled.
     * @param        height the height of the arc to be Filled.
     * @param        startAngle the beginning angle.
     * @param        arcAngle the angular extent of the arc,
     *                    relative to the start angle.
     * @see         java.awt.Graphics#drawArc
     */
    public void FillArc(int x, int y, int width, int height,
                        int startAngle, int arcAngle){
        Arc2D arc = new Arc2D.Float(x, y, width, height, startAngle, arcAngle, Arc2D.PIE);
        Fill(arc);
    }

    /**
     * Draws the outline of a circular or elliptical arc
     * covering the specified rectangle.
     * <p>
     * The resulting arc begins at <code>startAngle</code> and extends
     * for <code>arcAngle</code> degrees, using the current color.
     * Angles are interpreted such that 0&nbsp;degrees
     * is at the 3&nbsp;o'clock position.
     * A positive value indicates a counter-clockwise rotation
     * while a negative value indicates a clockwise rotation.
     * <p>
     * The center of the arc is the center of the rectangle whose origin
     * is (<i>x</i>,&nbsp;<i>y</i>) and whose size is specified by the
     * <code>width</code> and <code>height</code> arguments.
     * <p>
     * The resulting arc covers an area
     * <code>width&nbsp;+&nbsp;1</code> pixels wide
     * by <code>height&nbsp;+&nbsp;1</code> pixels tall.
     * <p>
     * The angles are specified relative to the non-square extents of
     * the bounding rectangle such that 45 degrees always falls on the
     * line from the center of the ellipse to the upper right corner of
     * the bounding rectangle. As a result, if the bounding rectangle is
     * noticeably longer in one axis than the other, the angles to the
     * start and end of the arc segment will be skewed farther along the
     * longer axis of the bounds.
     * @param        x the <i>x</i> coordinate of the
     *                    upper-left corner of the arc to be Drawn.
     * @param        y the <i>y</i>  coordinate of the
     *                    upper-left corner of the arc to be Drawn.
     * @param        width the width of the arc to be Drawn.
     * @param        height the height of the arc to be Drawn.
     * @param        startAngle the beginning angle.
     * @param        arcAngle the angular extent of the arc,
     *                    relative to the start angle.
     * @see         java.awt.Graphics#FillArc
     */
    public void DrawArc(int x, int y, int width, int height,
                        int startAngle, int arcAngle) {
        Arc2D arc = new Arc2D.Float(x, y, width, height, startAngle, arcAngle, Arc2D.OPEN);
        Draw(arc);
    }


    /**
     * Draws a sequence of connected lines defined by
     * arrays of <i>x</i> and <i>y</i> coordinates.
     * Each pair of (<i>x</i>,&nbsp;<i>y</i>) coordinates defines a point.
     * The figure is not closed if the first point
     * differs from the last point.
     * @param       xPoints an array of <i>x</i> points
     * @param       yPoints an array of <i>y</i> points
     * @param       nPoints the total number of points
     * @see         java.awt.Graphics#drawPolygon(int[], int[], int)
     * @since       JDK1.1
     */
    public void DrawPolyline(int[] xPoints, int[] yPoints,
                             int nPoints){
        if(nPoints > 0){
            GeneralPath path = new GeneralPath();
            path.moveTo(xPoints[0], yPoints[0]);
            for(int i=1; i<nPoints; i++)
                path.lineTo(xPoints[i], yPoints[i]);

            Draw(path);
        }
    }

    /**
     * Draws the outline of an oval.
     * The result is a circle or ellipse that fits within the
     * rectangle specified by the <code>x</code>, <code>y</code>,
     * <code>width</code>, and <code>height</code> arguments.
     * <p>
     * The oval covers an area that is
     * <code>width&nbsp;+&nbsp;1</code> pixels wide
     * and <code>height&nbsp;+&nbsp;1</code> pixels tall.
     * @param       x the <i>x</i> coordinate of the upper left
     *                     corner of the oval to be Drawn.
     * @param       y the <i>y</i> coordinate of the upper left
     *                     corner of the oval to be Drawn.
     * @param       width the width of the oval to be Drawn.
     * @param       height the height of the oval to be Drawn.
     * @see         java.awt.Graphics#FillOval
     */
    public void DrawOval(int x, int y, int width, int height){
        Ellipse2D oval = new Ellipse2D.Float(x, y, width, height);
        Draw(oval);
    }

    /**
     * Draws as much of the specified image as is currently available.
     * The image is Drawn with its top-left corner at
     * (<i>x</i>,&nbsp;<i>y</i>) in this graphics context's coordinate
     * space.  Transparent pixels are Drawn in the specified
     * background color.
     * <p>
     * This operation is equivalent to Filling a rectangle of the
     * width and height of the specified image with the given color and then
     * Drawing the image on top of it, but possibly more efficient.
     * <p>
     * This method returns immediately in all cases, even if the
     * complete image has not yet been loaded, and it has not been dithered
     * and Converted for the current output device.
     * <p>
     * If the image has not yet been completely loaded, then
     * <code>drawImage</code> returns <code>false</code>. As more of
     * the image becomes available, the process that Draws the image notifies
     * the specified image observer.
     * @param    img    the specified image to be Drawn.
     * @param    x      the <i>x</i> coordinate.
     * @param    y      the <i>y</i> coordinate.
     * @param    bgcolor the background color to paint under the
     *                         non-opaque portions of the image.
     * @param    observer    object to be notified as more of
     *                          the image is Converted.
     * @see      java.awt.Image
     * @see      java.awt.image.ImageObserver
     * @see      java.awt.image.ImageObserver#imageUpdate(java.awt.Image, int, int, int, int, int)
     */
    public bool DrawImage(Image img, int x, int y,
                             Color bgcolor,
                             ImageObserver observer){
        log.log(POILogger.WARN, "Not implemented");

        return false;
    }

    /**
     * Draws as much of the specified image as has already been scaled
     * to fit inside the specified rectangle.
     * <p>
     * The image is Drawn inside the specified rectangle of this
     * graphics context's coordinate space, and is scaled if
     * necessary. Transparent pixels are Drawn in the specified
     * background color.
     * This operation is equivalent to Filling a rectangle of the
     * width and height of the specified image with the given color and then
     * Drawing the image on top of it, but possibly more efficient.
     * <p>
     * This method returns immediately in all cases, even if the
     * entire image has not yet been scaled, dithered, and Converted
     * for the current output device.
     * If the current output representation is not yet complete then
     * <code>drawImage</code> returns <code>false</code>. As more of
     * the image becomes available, the process that Draws the image notifies
     * the specified image observer.
     * <p>
     * A scaled version of an image will not necessarily be
     * available immediately just because an unscaled version of the
     * image has been constructed for this output device.  Each size of
     * the image may be cached Separately and generated from the original
     * data in a separate image production sequence.
     * @param    img       the specified image to be Drawn.
     * @param    x         the <i>x</i> coordinate.
     * @param    y         the <i>y</i> coordinate.
     * @param    width     the width of the rectangle.
     * @param    height    the height of the rectangle.
     * @param    bgcolor   the background color to paint under the
     *                         non-opaque portions of the image.
     * @param    observer    object to be notified as more of
     *                          the image is Converted.
     * @see      java.awt.Image
     * @see      java.awt.image.ImageObserver
     * @see      java.awt.image.ImageObserver#imageUpdate(java.awt.Image, int, int, int, int, int)
     */
    public bool DrawImage(Image img, int x, int y,
                             int width, int height,
                             Color bgcolor,
                             ImageObserver observer){
        log.log(POILogger.WARN, "Not implemented");

        return false;
    }


    /**
     * Draws as much of the specified area of the specified image as is
     * currently available, scaling it on the fly to fit inside the
     * specified area of the destination Drawable surface. Transparent pixels
     * do not affect whatever pixels are already there.
     * <p>
     * This method returns immediately in all cases, even if the
     * image area to be Drawn has not yet been scaled, dithered, and Converted
     * for the current output device.
     * If the current output representation is not yet complete then
     * <code>drawImage</code> returns <code>false</code>. As more of
     * the image becomes available, the process that Draws the image notifies
     * the specified image observer.
     * <p>
     * This method always uses the unscaled version of the image
     * to render the scaled rectangle and performs the required
     * scaling on the fly. It does not use a cached, scaled version
     * of the image for this operation. Scaling of the image from source
     * to destination is performed such that the first coordinate
     * of the source rectangle is mapped to the first coordinate of
     * the destination rectangle, and the second source coordinate is
     * mapped to the second destination coordinate. The subimage is
     * scaled and flipped as needed to preserve those mappings.
     * @param       img the specified image to be Drawn
     * @param       dx1 the <i>x</i> coordinate of the first corner of the
     *                    destination rectangle.
     * @param       dy1 the <i>y</i> coordinate of the first corner of the
     *                    destination rectangle.
     * @param       dx2 the <i>x</i> coordinate of the second corner of the
     *                    destination rectangle.
     * @param       dy2 the <i>y</i> coordinate of the second corner of the
     *                    destination rectangle.
     * @param       sx1 the <i>x</i> coordinate of the first corner of the
     *                    source rectangle.
     * @param       sy1 the <i>y</i> coordinate of the first corner of the
     *                    source rectangle.
     * @param       sx2 the <i>x</i> coordinate of the second corner of the
     *                    source rectangle.
     * @param       sy2 the <i>y</i> coordinate of the second corner of the
     *                    source rectangle.
     * @param       observer object to be notified as more of the image is
     *                    scaled and Converted.
     * @see         java.awt.Image
     * @see         java.awt.image.ImageObserver
     * @see         java.awt.image.ImageObserver#imageUpdate(java.awt.Image, int, int, int, int, int)
     * @since       JDK1.1
     */
    public bool DrawImage(Image img,
                             int dx1, int dy1, int dx2, int dy2,
                             int sx1, int sy1, int sx2, int sy2,
                             ImageObserver observer){
        log.log(POILogger.WARN, "Not implemented");
        return false;
    }

    /**
     * Draws as much of the specified area of the specified image as is
     * currently available, scaling it on the fly to fit inside the
     * specified area of the destination Drawable surface.
     * <p>
     * Transparent pixels are Drawn in the specified background color.
     * This operation is equivalent to Filling a rectangle of the
     * width and height of the specified image with the given color and then
     * Drawing the image on top of it, but possibly more efficient.
     * <p>
     * This method returns immediately in all cases, even if the
     * image area to be Drawn has not yet been scaled, dithered, and Converted
     * for the current output device.
     * If the current output representation is not yet complete then
     * <code>drawImage</code> returns <code>false</code>. As more of
     * the image becomes available, the process that Draws the image notifies
     * the specified image observer.
     * <p>
     * This method always uses the unscaled version of the image
     * to render the scaled rectangle and performs the required
     * scaling on the fly. It does not use a cached, scaled version
     * of the image for this operation. Scaling of the image from source
     * to destination is performed such that the first coordinate
     * of the source rectangle is mapped to the first coordinate of
     * the destination rectangle, and the second source coordinate is
     * mapped to the second destination coordinate. The subimage is
     * scaled and flipped as needed to preserve those mappings.
     * @param       img the specified image to be Drawn
     * @param       dx1 the <i>x</i> coordinate of the first corner of the
     *                    destination rectangle.
     * @param       dy1 the <i>y</i> coordinate of the first corner of the
     *                    destination rectangle.
     * @param       dx2 the <i>x</i> coordinate of the second corner of the
     *                    destination rectangle.
     * @param       dy2 the <i>y</i> coordinate of the second corner of the
     *                    destination rectangle.
     * @param       sx1 the <i>x</i> coordinate of the first corner of the
     *                    source rectangle.
     * @param       sy1 the <i>y</i> coordinate of the first corner of the
     *                    source rectangle.
     * @param       sx2 the <i>x</i> coordinate of the second corner of the
     *                    source rectangle.
     * @param       sy2 the <i>y</i> coordinate of the second corner of the
     *                    source rectangle.
     * @param       bgcolor the background color to paint under the
     *                    non-opaque portions of the image.
     * @param       observer object to be notified as more of the image is
     *                    scaled and Converted.
     * @see         java.awt.Image
     * @see         java.awt.image.ImageObserver
     * @see         java.awt.image.ImageObserver#imageUpdate(java.awt.Image, int, int, int, int, int)
     * @since       JDK1.1
     */
    public bool DrawImage(Image img,
                             int dx1, int dy1, int dx2, int dy2,
                             int sx1, int sy1, int sx2, int sy2,
                             Color bgcolor,
                             ImageObserver observer){
        log.log(POILogger.WARN, "Not implemented");
        return false;
    }

    /**
     * Draws as much of the specified image as is currently available.
     * The image is Drawn with its top-left corner at
     * (<i>x</i>,&nbsp;<i>y</i>) in this graphics context's coordinate
     * space. Transparent pixels in the image do not affect whatever
     * pixels are already there.
     * <p>
     * This method returns immediately in all cases, even if the
     * complete image has not yet been loaded, and it has not been dithered
     * and Converted for the current output device.
     * <p>
     * If the image has completely loaded and its pixels are
     * no longer being Changed, then
     * <code>drawImage</code> returns <code>true</code>.
     * Otherwise, <code>drawImage</code> returns <code>false</code>
     * and as more of
     * the image becomes available
     * or it is time to Draw another frame of animation,
     * the process that loads the image notifies
     * the specified image observer.
     * @param    img the specified image to be Drawn. This method does
     *               nothing if <code>img</code> is null.
     * @param    x   the <i>x</i> coordinate.
     * @param    y   the <i>y</i> coordinate.
     * @param    observer    object to be notified as more of
     *                          the image is Converted.
     * @return   <code>false</code> if the image pixels are still changing;
     *           <code>true</code> otherwise.
     * @see      java.awt.Image
     * @see      java.awt.image.ImageObserver
     * @see      java.awt.image.ImageObserver#imageUpdate(java.awt.Image, int, int, int, int, int)
     */
    public bool DrawImage(Image img, int x, int y,
                             ImageObserver observer) {
        log.log(POILogger.WARN, "Not implemented");
        return false;
    }

    /**
     * Disposes of this graphics context and releases
     * any system resources that it is using.
     * A <code>Graphics</code> object cannot be used after
     * <code>dispose</code>has been called.
     * <p>
     * When a Java program Runs, a large number of <code>Graphics</code>
     * objects can be Created within a short time frame.
     * Although the finalization process of the garbage collector
     * also disposes of the same system resources, it is preferable
     * to manually free the associated resources by calling this
     * method rather than to rely on a finalization process which
     * may not run to completion for a long period of time.
     * <p>
     * Graphics objects which are provided as arguments to the
     * <code>paint</code> and <code>update</code> methods
     * of components are automatically released by the system when
     * those methods return. For efficiency, programmers should
     * call <code>dispose</code> when finished using
     * a <code>Graphics</code> object only if it was Created
     * directly from a component or another <code>Graphics</code> object.
     * @see         java.awt.Graphics#finalize
     * @see         java.awt.Component#paint
     * @see         java.awt.Component#update
     * @see         java.awt.Component#getGraphics
     * @see         java.awt.Graphics#create
     */
    public void dispose() {
        ;
    }

    /**
     * Draws a line, using the current color, between the points
     * <code>(x1,&nbsp;y1)</code> and <code>(x2,&nbsp;y2)</code>
     * in this graphics context's coordinate system.
     * @param   x1  the first point's <i>x</i> coordinate.
     * @param   y1  the first point's <i>y</i> coordinate.
     * @param   x2  the second point's <i>x</i> coordinate.
     * @param   y2  the second point's <i>y</i> coordinate.
     */
    public void DrawLine(int x1, int y1, int x2, int y2){
        Line2D line = new Line2D.Float(x1, y1, x2, y2);
        Draw(line);
    }

    /**
     * Fills a closed polygon defined by
     * arrays of <i>x</i> and <i>y</i> coordinates.
     * <p>
     * This method Draws the polygon defined by <code>nPoint</code> line
     * segments, where the first <code>nPoint&nbsp;-&nbsp;1</code>
     * line segments are line segments from
     * <code>(xPoints[i&nbsp;-&nbsp;1],&nbsp;yPoints[i&nbsp;-&nbsp;1])</code>
     * to <code>(xPoints[i],&nbsp;yPoints[i])</code>, for
     * 1&nbsp;&le;&nbsp;<i>i</i>&nbsp;&le;&nbsp;<code>nPoints</code>.
     * The figure is automatically closed by Drawing a line connecting
     * the point to the first point, if those points are different.
     * <p>
     * The area inside the polygon is defined using an
     * even-odd fill rule, also known as the alternating rule.
     * @param        xPoints   a an array of <code>x</code> coordinates.
     * @param        yPoints   a an array of <code>y</code> coordinates.
     * @param        nPoints   a the total number of points.
     * @see          java.awt.Graphics#drawPolygon(int[], int[], int)
     */
    public void FillPolygon(int[] xPoints, int[] yPoints,
                            int nPoints){
        java.awt.Polygon polygon = new java.awt.Polygon(xPoints, yPoints, nPoints);
        Fill(polygon);
    }

    /**
     * Fills the specified rectangle.
     * The left and right edges of the rectangle are at
     * <code>x</code> and <code>x&nbsp;+&nbsp;width&nbsp;-&nbsp;1</code>.
     * The top and bottom edges are at
     * <code>y</code> and <code>y&nbsp;+&nbsp;height&nbsp;-&nbsp;1</code>.
     * The resulting rectangle covers an area
     * <code>width</code> pixels wide by
     * <code>height</code> pixels tall.
     * The rectangle is Filled using the graphics context's current color.
     * @param         x   the <i>x</i> coordinate
     *                         of the rectangle to be Filled.
     * @param         y   the <i>y</i> coordinate
     *                         of the rectangle to be Filled.
     * @param         width   the width of the rectangle to be Filled.
     * @param         height   the height of the rectangle to be Filled.
     * @see           java.awt.Graphics#ClearRect
     * @see           java.awt.Graphics#drawRect
     */
    public void FillRect(int x, int y, int width, int height){
        Rectangle rect = new Rectangle(x, y, width, height);
        Fill(rect);
    }

    /**
     * Draws the outline of the specified rectangle.
     * The left and right edges of the rectangle are at
     * <code>x</code> and <code>x&nbsp;+&nbsp;width</code>.
     * The top and bottom edges are at
     * <code>y</code> and <code>y&nbsp;+&nbsp;height</code>.
     * The rectangle is Drawn using the graphics context's current color.
     * @param         x   the <i>x</i> coordinate
     *                         of the rectangle to be Drawn.
     * @param         y   the <i>y</i> coordinate
     *                         of the rectangle to be Drawn.
     * @param         width   the width of the rectangle to be Drawn.
     * @param         height   the height of the rectangle to be Drawn.
     * @see          java.awt.Graphics#FillRect
     * @see          java.awt.Graphics#ClearRect
     */
    public void DrawRect(int x, int y, int width, int height) {
        Rectangle rect = new Rectangle(x, y, width, height);
        Draw(rect);
    }

    /**
     * Draws a closed polygon defined by
     * arrays of <i>x</i> and <i>y</i> coordinates.
     * Each pair of (<i>x</i>,&nbsp;<i>y</i>) coordinates defines a point.
     * <p>
     * This method Draws the polygon defined by <code>nPoint</code> line
     * segments, where the first <code>nPoint&nbsp;-&nbsp;1</code>
     * line segments are line segments from
     * <code>(xPoints[i&nbsp;-&nbsp;1],&nbsp;yPoints[i&nbsp;-&nbsp;1])</code>
     * to <code>(xPoints[i],&nbsp;yPoints[i])</code>, for
     * 1&nbsp;&le;&nbsp;<i>i</i>&nbsp;&le;&nbsp;<code>nPoints</code>.
     * The figure is automatically closed by Drawing a line connecting
     * the point to the first point, if those points are different.
     * @param        xPoints   a an array of <code>x</code> coordinates.
     * @param        yPoints   a an array of <code>y</code> coordinates.
     * @param        nPoints   a the total number of points.
     * @see          java.awt.Graphics#FillPolygon(int[],int[],int)
     * @see          java.awt.Graphics#drawPolyline
     */
    public void DrawPolygon(int[] xPoints, int[] yPoints,
                            int nPoints){
        java.awt.Polygon polygon = new java.awt.Polygon(xPoints, yPoints, nPoints);
        Draw(polygon);
    }

    /**
     * Intersects the current clip with the specified rectangle.
     * The resulting clipping area is the intersection of the current
     * clipping area and the specified rectangle.  If there is no
     * current clipping area, either because the clip has never been
     * Set, or the clip has been Cleared using <code>setClip(null)</code>,
     * the specified rectangle becomes the new clip.
     * This method Sets the user clip, which is independent of the
     * clipping associated with device bounds and window visibility.
     * This method can only be used to make the current clip smaller.
     * To Set the current clip larger, use any of the SetClip methods.
     * Rendering operations have no effect outside of the clipping area.
     * @param x the x coordinate of the rectangle to intersect the clip with
     * @param y the y coordinate of the rectangle to intersect the clip with
     * @param width the width of the rectangle to intersect the clip with
     * @param height the height of the rectangle to intersect the clip with
     * @see #setClip(int, int, int, int)
     * @see #setClip(Shape)
     */
    public void clipRect(int x, int y, int width, int height){
        clip(new Rectangle(x, y, width, height));
    }

    /**
     * Sets the current clipping area to an arbitrary clip shape.
     * Not all objects that implement the <code>Shape</code>
     * interface can be used to Set the clip.  The only
     * <code>Shape</code> objects that are guaranteed to be
     * supported are <code>Shape</code> objects that are
     * obtained via the <code>getClip</code> method and via
     * <code>Rectangle</code> objects.  This method Sets the
     * user clip, which is independent of the clipping associated
     * with device bounds and window visibility.
     * @param clip the <code>Shape</code> to use to Set the clip
     * @see         java.awt.Graphics#getClip()
     * @see         java.awt.Graphics#clipRect
     * @see         java.awt.Graphics#setClip(int, int, int, int)
     * @since       JDK1.1
     */
    public void SetClip(Shape clip) {
        log.log(POILogger.WARN, "Not implemented");
    }

    /**
     * Returns the bounding rectangle of the current clipping area.
     * This method refers to the user clip, which is independent of the
     * clipping associated with device bounds and window visibility.
     * If no clip has previously been Set, or if the clip has been
     * Cleared using <code>setClip(null)</code>, this method returns
     * <code>null</code>.
     * The coordinates in the rectangle are relative to the coordinate
     * system origin of this graphics context.
     * @return      the bounding rectangle of the current clipping area,
     *              or <code>null</code> if no clip is Set.
     * @see         java.awt.Graphics#getClip
     * @see         java.awt.Graphics#clipRect
     * @see         java.awt.Graphics#setClip(int, int, int, int)
     * @see         java.awt.Graphics#setClip(Shape)
     * @since       JDK1.1
     */
    public Rectangle GetClipBounds(){
        Shape c = GetClip();
        if (c==null) {
            return null;
        }
        return c.GetBounds();
    }

    /**
     * Draws the text given by the specified iterator, using this
     * graphics context's current color. The iterator has to specify a font
     * for each character. The baseline of the
     * first character is at position (<i>x</i>,&nbsp;<i>y</i>) in this
     * graphics context's coordinate system.
     * @param       iterator the iterator whose text is to be Drawn
     * @param       x        the <i>x</i> coordinate.
     * @param       y        the <i>y</i> coordinate.
     * @see         java.awt.Graphics#drawBytes
     * @see         java.awt.Graphics#drawChars
     */
    public void DrawString(AttributedCharacterIterator iterator,
                           int x, int y){
        DrawString(iterator, (float)x, (float)y);
    }

    /**
     * Clears the specified rectangle by Filling it with the background
     * color of the current Drawing surface. This operation does not
     * use the current paint mode.
     * <p>
     * Beginning with Java&nbsp;1.1, the background color
     * of offscreen images may be system dependent. Applications should
     * use <code>setColor</code> followed by <code>FillRect</code> to
     * ensure that an offscreen image is Cleared to a specific color.
     * @param       x the <i>x</i> coordinate of the rectangle to Clear.
     * @param       y the <i>y</i> coordinate of the rectangle to Clear.
     * @param       width the width of the rectangle to Clear.
     * @param       height the height of the rectangle to Clear.
     * @see         java.awt.Graphics#FillRect(int, int, int, int)
     * @see         java.awt.Graphics#drawRect
     * @see         java.awt.Graphics#setColor(java.awt.Color)
     * @see         java.awt.Graphics#setPaintMode
     * @see         java.awt.Graphics#setXORMode(java.awt.Color)
     */
    public void ClearRect(int x, int y, int width, int height) {
        Paint paint = GetPaint();
        SetColor(getBackground());
        FillRect(x, y, width, height);
        SetPaint(paint);
    }

    public void copyArea(int x, int y, int width, int height, int dx, int dy) {
        ;
    }

    /**
     * Sets the current clip to the rectangle specified by the given
     * coordinates.  This method Sets the user clip, which is
     * independent of the clipping associated with device bounds
     * and window visibility.
     * Rendering operations have no effect outside of the clipping area.
     * @param       x the <i>x</i> coordinate of the new clip rectangle.
     * @param       y the <i>y</i> coordinate of the new clip rectangle.
     * @param       width the width of the new clip rectangle.
     * @param       height the height of the new clip rectangle.
     * @see         java.awt.Graphics#clipRect
     * @see         java.awt.Graphics#setClip(Shape)
     * @since       JDK1.1
     */
    public void SetClip(int x, int y, int width, int height){
        SetClip(new Rectangle(x, y, width, height));
    }

    /**
     * Concatenates the current <code>Graphics2D</code>
     * <code>Transform</code> with a rotation transform.
     * Subsequent rendering is rotated by the specified radians relative
     * to the previous origin.
     * This is equivalent to calling <code>transform(R)</code>, where R is an
     * <code>AffineTransform</code> represented by the following matrix:
     * <pre>
     *          [   cos(theta)    -sin(theta)    0   ]
     *          [   sin(theta)     cos(theta)    0   ]
     *          [       0              0         1   ]
     * </pre>
     * Rotating with a positive angle theta rotates points on the positive
     * x axis toward the positive y axis.
     * @param theta the angle of rotation in radians
     */
    public void rotate(double theta){
        _transform.rotate(theta);
    }

    /**
     * Concatenates the current <code>Graphics2D</code>
     * <code>Transform</code> with a translated rotation
     * transform.  Subsequent rendering is transformed by a transform
     * which is constructed by translating to the specified location,
     * rotating by the specified radians, and translating back by the same
     * amount as the original translation.  This is equivalent to the
     * following sequence of calls:
     * <pre>
     *          translate(x, y);
     *          rotate(theta);
     *          translate(-x, -y);
     * </pre>
     * Rotating with a positive angle theta rotates points on the positive
     * x axis toward the positive y axis.
     * @param theta the angle of rotation in radians
     * @param x x coordinate of the origin of the rotation
     * @param y y coordinate of the origin of the rotation
     */
    public void rotate(double theta, double x, double y){
        _transform.rotate(theta, x, y);
    }

    /**
     * Concatenates the current <code>Graphics2D</code>
     * <code>Transform</code> with a shearing transform.
     * Subsequent renderings are sheared by the specified
     * multiplier relative to the previous position.
     * This is equivalent to calling <code>transform(SH)</code>, where SH
     * is an <code>AffineTransform</code> represented by the following
     * matrix:
     * <pre>
     *          [   1   shx   0   ]
     *          [  shy   1    0   ]
     *          [   0    0    1   ]
     * </pre>
     * @param shx the multiplier by which coordinates are Shifted in
     * the positive X axis direction as a function of their Y coordinate
     * @param shy the multiplier by which coordinates are Shifted in
     * the positive Y axis direction as a function of their X coordinate
     */
    public void shear(double shx, double shy){
        _transform.shear(shx, shy);
    }

    /**
     * Get the rendering context of the <code>Font</code> within this
     * <code>Graphics2D</code> context.
     * The {@link FontRenderContext}
     * encapsulates application hints such as anti-aliasing and
     * fractional metrics, as well as target device specific information
     * such as dots-per-inch.  This information should be provided by the
     * application when using objects that perform typographical
     * formatting, such as <code>Font</code> and
     * <code>TextLayout</code>.  This information should also be provided
     * by applications that perform their own layout and need accurate
     * measurements of various characteristics of glyphs such as advance
     * and line height when various rendering hints have been applied to
     * the text rendering.
     *
     * @return a reference to an instance of FontRenderContext.
     * @see java.awt.font.FontRenderContext
     * @see java.awt.Font#CreateGlyphVector(FontRenderContext,char[])
     * @see java.awt.font.TextLayout
     * @since     JDK1.2
     */
    public FontRenderContext GetFontRenderContext() {
        bool IsAntiAliased = RenderingHints.VALUE_TEXT_ANTIALIAS_ON.Equals(
                GetRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING));
        bool usesFractionalMetrics = RenderingHints.VALUE_FRACTIONALMETRICS_ON.Equals(
                GetRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS));


        return new FontRenderContext(new AffineTransform(), isAntiAliased, usesFractionalMetrics);
    }

    /**
     * Composes an <code>AffineTransform</code> object with the
     * <code>Transform</code> in this <code>Graphics2D</code> according
     * to the rule last-specified-first-applied.  If the current
     * <code>Transform</code> is Cx, the result of composition
     * with Tx is a new <code>Transform</code> Cx'.  Cx' becomes the
     * current <code>Transform</code> for this <code>Graphics2D</code>.
     * Transforming a point p by the updated <code>Transform</code> Cx' is
     * equivalent to first transforming p by Tx and then transforming
     * the result by the original <code>Transform</code> Cx.  In other
     * words, Cx'(p) = Cx(Tx(p)).  A copy of the Tx is made, if necessary,
     * so further modifications to Tx do not affect rendering.
     * @param Tx the <code>AffineTransform</code> object to be composed with
     * the current <code>Transform</code>
     * @see #setTransform
     * @see AffineTransform
     */
    public void transform(AffineTransform Tx) {
        _transform.concatenate(Tx);
    }

    /**
     * Renders a <code>BufferedImage</code> that is
     * filtered with a
     * {@link BufferedImageOp}.
     * The rendering attributes applied include the <code>Clip</code>,
     * <code>Transform</code>
     * and <code>Composite</code> attributes.  This is equivalent to:
     * <pre>
     * img1 = op.filter(img, null);
     * DrawImage(img1, new AffineTransform(1f,0f,0f,1f,x,y), null);
     * </pre>
     * @param img the <code>BufferedImage</code> to be rendered
     * @param op the filter to be applied to the image before rendering
     * @param x the x coordinate in user space where the image is rendered
     * @param y the y coordinate in user space where the image is rendered
     * @see #_transform
     * @see #setTransform
     * @see #setComposite
     * @see #clip
     * @see #setClip(Shape)
     */
    public void DrawImage(BufferedImage img,
                          BufferedImageOp op,
                          int x,
                          int y){
        img = op.filter(img, null);
        DrawImage(img, x, y, null);
    }

    /**
     * Sets the background color for the <code>Graphics2D</code> context.
     * The background color is used for Clearing a region.
     * When a <code>Graphics2D</code> is constructed for a
     * <code>Component</code>, the background color is
     * inherited from the <code>Component</code>. Setting the background color
     * in the <code>Graphics2D</code> context only affects the subsequent
     * <code>ClearRect</code> calls and not the background color of the
     * <code>Component</code>.  To change the background
     * of the <code>Component</code>, use appropriate methods of
     * the <code>Component</code>.
     * @param color the background color that isused in
     * subsequent calls to <code>ClearRect</code>
     * @see #getBackground
     * @see java.awt.Graphics#ClearRect
     */
    public void SetBackground(Color color) {
        if(color == null)
            return;

        _background = color;
    }

    /**
     * Returns the background color used for Clearing a region.
     * @return the current <code>Graphics2D</code> <code>Color</code>,
     * which defines the background color.
     * @see #setBackground
     */
    public Color GetBackground(){
        return _background;
    }

    /**
     * Sets the <code>Composite</code> for the <code>Graphics2D</code> context.
     * The <code>Composite</code> is used in all Drawing methods such as
     * <code>drawImage</code>, <code>drawString</code>, <code>draw</code>,
     * and <code>Fill</code>.  It specifies how new pixels are to be combined
     * with the existing pixels on the graphics device during the rendering
     * Process.
     * <p>If this <code>Graphics2D</code> context is Drawing to a
     * <code>Component</code> on the display screen and the
     * <code>Composite</code> is a custom object rather than an
     * instance of the <code>AlphaComposite</code> class, and if
     * there is a security manager, its <code>CheckPermission</code>
     * method is called with an <code>AWTPermission("readDisplayPixels")</code>
     * permission.
     *
     * @param comp the <code>Composite</code> object to be used for rendering
     * @throws SecurityException
     *         if a custom <code>Composite</code> object is being
     *         used to render to the screen and a security manager
     *         is Set and its <code>CheckPermission</code> method
     *         does not allow the operation.
     * @see java.awt.Graphics#setXORMode
     * @see java.awt.Graphics#setPaintMode
     * @see java.awt.AlphaComposite
     */
    public void SetComposite(Composite comp){
        log.log(POILogger.WARN, "Not implemented");
    }

    /**
     * Returns the current <code>Composite</code> in the
     * <code>Graphics2D</code> context.
     * @return the current <code>Graphics2D</code> <code>Composite</code>,
     *              which defines a compositing style.
     * @see #setComposite
     */
    public Composite GetComposite(){
        log.log(POILogger.WARN, "Not implemented");
        return null;
    }

    /**
     * Returns the value of a single preference for the rendering algorithms.
     * Hint categories include controls for rendering quality and overall
     * time/quality trade-off in the rendering Process.  Refer to the
     * <code>RenderingHints</code> class for defInitions of some common
     * keys and values.
     * @param hintKey the key corresponding to the hint to Get.
     * @return an object representing the value for the specified hint key.
     * Some of the keys and their associated values are defined in the
     * <code>RenderingHints</code> class.
     * @see RenderingHints
     */
    public Object GetRenderingHint(RenderingHints.Key hintKey){
        return _hints.Get(hintKey);
    }

    /**
     * Sets the value of a single preference for the rendering algorithms.
     * Hint categories include controls for rendering quality and overall
     * time/quality trade-off in the rendering Process.  Refer to the
     * <code>RenderingHints</code> class for defInitions of some common
     * keys and values.
     * @param hintKey the key of the hint to be Set.
     * @param hintValue the value indicating preferences for the specified
     * hint category.
     * @see RenderingHints
     */
    public void SetRenderingHint(RenderingHints.Key hintKey, Object hintValue){
        _hints.Put(hintKey, hintValue);
    }


    /**
     * Renders the text of the specified
     * {@link GlyphVector} using
     * the <code>Graphics2D</code> context's rendering attributes.
     * The rendering attributes applied include the <code>Clip</code>,
     * <code>Transform</code>, <code>Paint</code>, and
     * <code>Composite</code> attributes.  The <code>GlyphVector</code>
     * specifies individual glyphs from a {@link Font}.
     * The <code>GlyphVector</code> can also contain the glyph positions.
     * This is the fastest way to render a Set of characters to the
     * screen.
     *
     * @param g the <code>GlyphVector</code> to be rendered
     * @param x the x position in user space where the glyphs should be
     *        rendered
     * @param y the y position in user space where the glyphs should be
     *        rendered
     *
     * @see java.awt.Font#CreateGlyphVector(FontRenderContext, char[])
     * @see java.awt.font.GlyphVector
     * @see #setPaint
     * @see java.awt.Graphics#setColor
     * @see #setTransform
     * @see #setComposite
     * @see #setClip(Shape)
     */
    public void DrawGlyphVector(GlyphVector g, float x, float y) {
        Shape glyphOutline = g.GetOutline(x, y);
        Fill(glyphOutline);
    }

    /**
     * Returns the device configuration associated with this
     * <code>Graphics2D</code>.
     * @return the device configuration
     */
    public GraphicsConfiguration GetDeviceConfiguration() {
        return GraphicsEnvironment.GetLocalGraphicsEnvironment().
                GetDefaultScreenDevice().GetDefaultConfiguration();
    }

    /**
     * Sets the values of an arbitrary number of preferences for the
     * rendering algorithms.
     * Only values for the rendering hints that are present in the
     * specified <code>Dictionary</code> object are modified.
     * All other preferences not present in the specified
     * object are left unmodified.
     * Hint categories include controls for rendering quality and
     * overall time/quality trade-off in the rendering Process.
     * Refer to the <code>RenderingHints</code> class for defInitions of
     * some common keys and values.
     * @param hints the rendering hints to be Set
     * @see RenderingHints
     */
    public void AddRenderingHints(Map hints){
        this._hints.PutAll(hints);
    }

    /**
     * Concatenates the current
     * <code>Graphics2D</code> <code>Transform</code>
     * with a translation transform.
     * Subsequent rendering is translated by the specified
     * distance relative to the previous position.
     * This is equivalent to calling transform(T), where T is an
     * <code>AffineTransform</code> represented by the following matrix:
     * <pre>
     *          [   1    0    tx  ]
     *          [   0    1    ty  ]
     *          [   0    0    1   ]
     * </pre>
     * @param tx the distance to translate along the x-axis
     * @param ty the distance to translate along the y-axis
     */
    public void translate(double tx, double ty){
        _transform.translate(tx, ty);
    }

    /**
     * Renders the text of the specified iterator, using the
     * <code>Graphics2D</code> context's current <code>Paint</code>. The
     * iterator must specify a font
     * for each character. The baseline of the
     * first character is at position (<i>x</i>,&nbsp;<i>y</i>) in the
     * User Space.
     * The rendering attributes applied include the <code>Clip</code>,
     * <code>Transform</code>, <code>Paint</code>, and
     * <code>Composite</code> attributes.
     * For characters in script systems such as Hebrew and Arabic,
     * the glyphs can be rendered from right to left, in which case the
     * coordinate supplied is the location of the leftmost character
     * on the baseline.
     * @param iterator the iterator whose text is to be rendered
     * @param x the x coordinate where the iterator's text is to be
     * rendered
     * @param y the y coordinate where the iterator's text is to be
     * rendered
     * @see #setPaint
     * @see java.awt.Graphics#setColor
     * @see #setTransform
     * @see #setComposite
     * @see #setClip
     */
    public void DrawString(AttributedCharacterIterator iterator, float x, float y) {
        log.log(POILogger.WARN, "Not implemented");
    }

    /**
     * Checks whether or not the specified <code>Shape</code> intersects
     * the specified {@link Rectangle}, which is in device
     * space. If <code>onStroke</code> is false, this method Checks
     * whether or not the interior of the specified <code>Shape</code>
     * intersects the specified <code>Rectangle</code>.  If
     * <code>onStroke</code> is <code>true</code>, this method Checks
     * whether or not the <code>Stroke</code> of the specified
     * <code>Shape</code> outline intersects the specified
     * <code>Rectangle</code>.
     * The rendering attributes taken into account include the
     * <code>Clip</code>, <code>Transform</code>, and <code>Stroke</code>
     * attributes.
     * @param rect the area in device space to check for a hit
     * @param s the <code>Shape</code> to check for a hit
     * @param onStroke flag used to choose between Testing the
     * stroked or the Filled shape.  If the flag is <code>true</code>, the
     * <code>Stroke</code> oultine is Tested.  If the flag is
     * <code>false</code>, the Filled <code>Shape</code> is Tested.
     * @return <code>true</code> if there is a hit; <code>false</code>
     * otherwise.
     * @see #setStroke
     * @see #Fill(Shape)
     * @see #draw(Shape)
     * @see #_transform
     * @see #setTransform
     * @see #clip
     * @see #setClip(Shape)
     */
    public bool hit(Rectangle rect,
                       Shape s,
                       bool onStroke){
        if (onStroke) {
            s = GetStroke().CreateStrokedShape(s);
        }

        s = GetTransform().CreateTransformedShape(s);

        return s.intersects(rect);
    }

    /**
     * Gets the preferences for the rendering algorithms.  Hint categories
     * include controls for rendering quality and overall time/quality
     * trade-off in the rendering Process.
     * Returns all of the hint key/value pairs that were ever specified in
     * one operation.  Refer to the
     * <code>RenderingHints</code> class for defInitions of some common
     * keys and values.
     * @return a reference to an instance of <code>RenderingHints</code>
     * that Contains the current preferences.
     * @see RenderingHints
     */
    public RenderingHints GetRenderingHints(){
        return _hints;
    }

    /**
     * Replaces the values of all preferences for the rendering
     * algorithms with the specified <code>hints</code>.
     * The existing values for all rendering hints are discarded and
     * the new Set of known hints and values are Initialized from the
     * specified {@link Map} object.
     * Hint categories include controls for rendering quality and
     * overall time/quality trade-off in the rendering Process.
     * Refer to the <code>RenderingHints</code> class for defInitions of
     * some common keys and values.
     * @param hints the rendering hints to be Set
     * @see RenderingHints
     */
    public void SetRenderingHints(Map hints){
        this._hints = new RenderingHints(hints);
    }

    /**
     * Renders an image, Applying a transform from image space into user space
     * before Drawing.
     * The transformation from user space into device space is done with
     * the current <code>Transform</code> in the <code>Graphics2D</code>.
     * The specified transformation is applied to the image before the
     * transform attribute in the <code>Graphics2D</code> context is applied.
     * The rendering attributes applied include the <code>Clip</code>,
     * <code>Transform</code>, and <code>Composite</code> attributes.
     * Note that no rendering is done if the specified transform is
     * noninvertible.
     * @param img the <code>Image</code> to be rendered
     * @param xform the transformation from image space into user space
     * @param obs the {@link ImageObserver}
     * to be notified as more of the <code>Image</code>
     * is Converted
     * @return <code>true</code> if the <code>Image</code> is
     * fully loaded and completely rendered;
     * <code>false</code> if the <code>Image</code> is still being loaded.
     * @see #_transform
     * @see #setTransform
     * @see #setComposite
     * @see #clip
     * @see #setClip(Shape)
     */
     public bool DrawImage(Image img, AffineTransform xform, ImageObserver obs) {
        log.log(POILogger.WARN, "Not implemented");
        return false;
    }

    /**
     * Draws as much of the specified image as has already been scaled
     * to fit inside the specified rectangle.
     * <p>
     * The image is Drawn inside the specified rectangle of this
     * graphics context's coordinate space, and is scaled if
     * necessary. Transparent pixels do not affect whatever pixels
     * are already there.
     * <p>
     * This method returns immediately in all cases, even if the
     * entire image has not yet been scaled, dithered, and Converted
     * for the current output device.
     * If the current output representation is not yet complete, then
     * <code>drawImage</code> returns <code>false</code>. As more of
     * the image becomes available, the process that loads the image notifies
     * the image observer by calling its <code>imageUpdate</code> method.
     * <p>
     * A scaled version of an image will not necessarily be
     * available immediately just because an unscaled version of the
     * image has been constructed for this output device.  Each size of
     * the image may be cached Separately and generated from the original
     * data in a separate image production sequence.
     * @param    img    the specified image to be Drawn. This method does
     *                  nothing if <code>img</code> is null.
     * @param    x      the <i>x</i> coordinate.
     * @param    y      the <i>y</i> coordinate.
     * @param    width  the width of the rectangle.
     * @param    height the height of the rectangle.
     * @param    observer    object to be notified as more of
     *                          the image is Converted.
     * @return   <code>false</code> if the image pixels are still changing;
     *           <code>true</code> otherwise.
     * @see      java.awt.Image
     * @see      java.awt.image.ImageObserver
     * @see      java.awt.image.ImageObserver#imageUpdate(java.awt.Image, int, int, int, int, int)
     */
    public bool DrawImage(Image img, int x, int y,
                             int width, int height,
                             ImageObserver observer) {
        log.log(POILogger.WARN, "Not implemented");
        return false;
    }

    /**
     * Creates a new <code>Graphics</code> object that is
     * a copy of this <code>Graphics</code> object.
     * @return     a new graphics context that is a copy of
     *                       this graphics context.
     */
    public Graphics Create() {
        try {
            return (Graphics)Clone();
        } catch (CloneNotSupportedException e){
            throw new HSLFException(e);
        }
    }

    /**
     * Gets the font metrics for the specified font.
     * @return    the font metrics for the specified font.
     * @param     f the specified font
     * @see       java.awt.Graphics#getFont
     * @see       java.awt.FontMetrics
     * @see       java.awt.Graphics#getFontMetrics()
     */
    @SuppressWarnings("deprecation")
    public FontMetrics GetFontMetrics(Font f) {
        return Toolkit.GetDefaultToolkit().GetFontMetrics(f);
    }

    /**
     * Sets the paint mode of this graphics context to alternate between
     * this graphics context's current color and the new specified color.
     * This specifies that logical pixel operations are performed in the
     * XOR mode, which alternates pixels between the current color and
     * a specified XOR color.
     * <p>
     * When Drawing operations are performed, pixels which are the
     * current color are Changed to the specified color, and vice versa.
     * <p>
     * Pixels that are of colors other than those two colors are Changed
     * in an unpredictable but reversible manner; if the same figure is
     * Drawn twice, then all pixels are restored to their original values.
     * @param     c1 the XOR alternation color
     */
    public void SetXORMode(Color c1) {
        log.log(POILogger.WARN, "Not implemented");
    }

    /**
     * Sets the paint mode of this graphics context to overwrite the
     * destination with this graphics context's current color.
     * This Sets the logical pixel operation function to the paint or
     * overwrite mode.  All subsequent rendering operations will
     * overwrite the destination with the current color.
     */
    public void SetPaintMode() {
        log.log(POILogger.WARN, "Not implemented");
    }

    /**
     * Renders a
     * {@link RenderableImage},
     * Applying a transform from image space into user space before Drawing.
     * The transformation from user space into device space is done with
     * the current <code>Transform</code> in the <code>Graphics2D</code>.
     * The specified transformation is applied to the image before the
     * transform attribute in the <code>Graphics2D</code> context is applied.
     * The rendering attributes applied include the <code>Clip</code>,
     * <code>Transform</code>, and <code>Composite</code> attributes. Note
     * that no rendering is done if the specified transform is
     * noninvertible.
     *<p>
     * Rendering hints Set on the <code>Graphics2D</code> object might
     * be used in rendering the <code>RenderableImage</code>.
     * If explicit control is required over specific hints recognized by a
     * specific <code>RenderableImage</code>, or if knowledge of which hints
     * are used is required, then a <code>RenderedImage</code> should be
     * obtained directly from the <code>RenderableImage</code>
     * and rendered using
     *{@link #drawRenderedImage(RenderedImage, AffineTransform) DrawRenderedImage}.
     * @param img the image to be rendered. This method does
     *            nothing if <code>img</code> is null.
     * @param xform the transformation from image space into user space
     * @see #_transform
     * @see #setTransform
     * @see #setComposite
     * @see #clip
     * @see #setClip
     * @see #drawRenderedImage
     */
     public void DrawRenderedImage(RenderedImage img, AffineTransform xform) {
        log.log(POILogger.WARN, "Not implemented");
    }

    /**
     * Renders a {@link RenderedImage},
     * Applying a transform from image
     * space into user space before Drawing.
     * The transformation from user space into device space is done with
     * the current <code>Transform</code> in the <code>Graphics2D</code>.
     * The specified transformation is applied to the image before the
     * transform attribute in the <code>Graphics2D</code> context is applied.
     * The rendering attributes applied include the <code>Clip</code>,
     * <code>Transform</code>, and <code>Composite</code> attributes. Note
     * that no rendering is done if the specified transform is
     * noninvertible.
     * @param img the image to be rendered. This method does
     *            nothing if <code>img</code> is null.
     * @param xform the transformation from image space into user space
     * @see #_transform
     * @see #setTransform
     * @see #setComposite
     * @see #clip
     * @see #setClip
     */
    public void DrawRenderableImage(RenderableImage img, AffineTransform xform) {
        log.log(POILogger.WARN, "Not implemented");
    }

    protected void ApplyStroke(SimpleShape shape) {
        if (_stroke is BasicStroke){
            BasicStroke bs = (BasicStroke)_stroke;
            shape.SetLineWidth(bs.GetLineWidth());
            float[] dash = bs.GetDashArray();
            if (dash != null) {
                //TODO: implement more dashing styles
                shape.SetLineDashing(Line.PEN_DASH);
            }
        }
    }

    protected void ApplyPaint(SimpleShape shape) {
        if (_paint is Color) {
            shape.GetFill().SetForegroundColor((Color)_paint);
        }
    }
}





