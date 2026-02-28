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















using NPOI.HSLF.record.TextRulerAtom;
using NPOI.HSLF.usermodel.RichTextRun;
using NPOI.util.POILogFactory;
using NPOI.util.POILogger;

/**
 * Paint text into java.awt.Graphics2D
 *
 * @author Yegor Kozlov
 */
public class TextPainter {
    protected POILogger logger = POILogFactory.GetLogger(this.GetClass());

    /**
     * Display unicode square if a bullet char can't be displayed,
     * for example, if Wingdings font is used.
     * TODO: map Wingdngs and Symbol to unicode Arial
     */
    protected static char DEFAULT_BULLET_CHAR = '\u25a0';

    protected TextShape _shape;

    public TextPainter(TextShape shape){
        _shape = shape;
    }

    /**
     * Convert the underlying Set of rich text Runs into java.text.AttributedString
     */
    public AttributedString GetAttributedString(TextRun txRun){
        String text = txRun.GetText();
        //TODO: properly process tabs
        text = text.Replace('\t', ' ');
        text = text.Replace((char)160, ' ');

        AttributedString at = new AttributedString(text);
        RichTextRun[] rt = txRun.GetRichTextRuns();
        for (int i = 0; i < rt.Length; i++) {
            int start = rt[i].GetStartIndex();
            int end = rt[i].GetEndIndex();
            if(start == end) {
                logger.log(POILogger.INFO,  "Skipping RichTextRun with zero length");
                continue;
            }

            at.AddAttribute(TextAttribute.FAMILY, rt[i].GetFontName(), start, end);
            at.AddAttribute(TextAttribute.SIZE, new Float(rt[i].GetFontSize()), start, end);
            at.AddAttribute(TextAttribute.FOREGROUND, rt[i].GetFontColor(), start, end);
            if(rt[i].IsBold()) at.AddAttribute(TextAttribute.WEIGHT, TextAttribute.WEIGHT_BOLD, start, end);
            if(rt[i].IsItalic()) at.AddAttribute(TextAttribute.POSTURE, TextAttribute.POSTURE_OBLIQUE, start, end);
            if(rt[i].IsUnderlined()) {
                at.AddAttribute(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON, start, end);
                at.AddAttribute(TextAttribute.INPUT_METHOD_UNDERLINE, TextAttribute.UNDERLINE_LOW_TWO_PIXEL, start, end);
            }
            if(rt[i].IsStrikethrough()) at.AddAttribute(TextAttribute.STRIKETHROUGH, TextAttribute.STRIKETHROUGH_ON, start, end);
            int superScript = rt[i].GetSuperscript();
            if(superScript != 0) at.AddAttribute(TextAttribute.SUPERSCRIPT, superScript > 0 ? TextAttribute.SUPERSCRIPT_SUPER : TextAttribute.SUPERSCRIPT_SUB, start, end);

        }
        return at;
    }

    public void paint(Graphics2D graphics){
        Rectangle2D anchor = _shape.GetLogicalAnchor2D();
        TextElement[] elem = GetTextElements((float)anchor.Width, graphics.GetFontRenderContext());
        if(elem == null) return;

        float textHeight = 0;
        for (int i = 0; i < elem.Length; i++) {
            textHeight += elem[i].ascent + elem[i].descent;
        }

        int valign = _shape.GetVerticalAlignment();
        double y0 = anchor.GetY();
        switch (valign){
            case TextShape.AnchorTopBaseline:
            case TextShape.AnchorTop:
                y0 += _shape.GetMarginTop();
                break;
            case TextShape.AnchorBottom:
                y0 += anchor.Height - textHeight - _shape.GetMarginBottom();
                break;
            default:
            case TextShape.AnchorMiddle:
                float delta =  (float)anchor.Height - textHeight - _shape.GetMarginTop() - _shape.GetMarginBottom();
                y0 += _shape.GetMarginTop()  + delta/2;
                break;
        }

        //finally Draw the text fragments
        for (int i = 0; i < elem.Length; i++) {
            y0 += elem[i].ascent;

            Point2D.Double pen = new Point2D.Double();
            pen.y = y0;
            switch (elem[i]._align) {
                default:
                case TextShape.AlignLeft:
                    pen.x = anchor.GetX() + _shape.GetMarginLeft();
                    break;
                case TextShape.AlignCenter:
                    pen.x = anchor.GetX() + _shape.GetMarginLeft() +
                            (anchor.Width - elem[i].advance - _shape.GetMarginLeft() - _shape.GetMarginRight()) / 2;
                    break;
                case TextShape.AlignRight:
                    pen.x = anchor.GetX() + _shape.GetMarginLeft() +
                            (anchor.Width - elem[i].advance - _shape.GetMarginLeft() - _shape.GetMarginRight());
                    break;
            }
            if(elem[i]._bullet != null){
                graphics.DrawString(elem[i]._bullet.GetIterator(), (float)(pen.x + elem[i]._bulletOffset), (float)pen.y);
            }
            AttributedCharacterIterator chIt = elem[i]._text.GetIterator();
            if(chIt.GetEndIndex() > chIt.GetBeginIndex()) {
                graphics.DrawString(chIt, (float)(pen.x + elem[i]._textOffset), (float)pen.y);
            }
            y0 += elem[i].descent;
        }
    }

    public TextElement[] GetTextElements(float textWidth, FontRenderContext frc){
        TextRun run = _shape.GetTextRun();
        if (run == null) return null;

        String text = Run.GetText();
        if (text == null || text.Equals("")) return null;

        AttributedString at = GetAttributedString(Run);

        AttributedCharacterIterator it = at.GetIterator();
        int paragraphStart = it.GetBeginIndex();
        int paragraphEnd = it.GetEndIndex();

        List<TextElement> lines = new List<TextElement>();
        LineBreakMeasurer measurer = new LineBreakMeasurer(it, frc);
        measurer.SetPosition(paragraphStart);
        while (measurer.GetPosition() < paragraphEnd) {
            int startIndex = measurer.GetPosition();
            int nextBreak = text.IndexOf('\n', measurer.GetPosition() + 1);

            bool prStart = text[startIndex] == '\n';
            if(prStart) measurer.SetPosition(startIndex++);

            RichTextRun rt = Run.GetRichTextRunAt(startIndex == text.Length ? (startIndex-1) : startIndex);
            if(rt == null) {
                logger.log(POILogger.WARN,  "RichTextRun not found at pos" + startIndex + "; text.Length: " + text.Length);
                break;
            }

            float wrappingWidth = textWidth - _shape.GetMarginLeft() - _shape.GetMarginRight();
            int bulletOffset = rt.GetBulletOffset();
            int textOffset = rt.GetTextOffset();
            int indent = rt.GetIndentLevel();

            TextRulerAtom ruler = Run.GetTextRuler();
            if(ruler != null) {
                int bullet_val = ruler.GetBulletOffsets()[indent]*Shape.POINT_DPI/Shape.MASTER_DPI;
                int text_val = ruler.GetTextOffsets()[indent]*Shape.POINT_DPI/Shape.MASTER_DPI;
                if(bullet_val > text_val){
                    int a = bullet_val;
                    bullet_val = text_val;
                    text_val = a;
                }
                if(bullet_val != 0 ) bulletOffset = bullet_val;
                if(text_val != 0) textOffset = text_val;
            }

            if(bulletOffset > 0 || prStart || startIndex == 0) wrappingWidth -= textOffset;

            if (_shape.GetWordWrap() == TextShape.WrapNone) {
                wrappingWidth = _shape.Sheet.GetSlideShow().GetPageSize().width;
            }

            TextLayout textLayout = measurer.nextLayout(wrappingWidth + 1,
                    nextBreak == -1 ? paragraphEnd : nextBreak, true);
            if (textLayout == null) {
                textLayout = measurer.nextLayout(textWidth,
                    nextBreak == -1 ? paragraphEnd : nextBreak, false);
            }
            if(textLayout == null){
                logger.log(POILogger.WARN, "Failed to break text into lines: wrappingWidth: "+wrappingWidth+
                        "; text: " + rt.GetText());
                measurer.SetPosition(rt.GetEndIndex());
                continue;
            }
            int endIndex = measurer.GetPosition();

            float lineHeight = (float)textLayout.GetBounds().Height;
            int linespacing = rt.GetLineSpacing();
            if(linespacing == 0) linespacing = 100;

            TextElement el = new TextElement();
            if(linespacing >= 0){
                el.ascent = textLayout.GetAscent()*linespacing/100;
            } else {
                el.ascent = -linespacing*Shape.POINT_DPI/Shape.MASTER_DPI;
            }

            el._align = rt.GetAlignment();
            el.advance = textLayout.GetAdvance();
            el._textOffset = textOffset;
            el._text = new AttributedString(it, startIndex, endIndex);
            el.textStartIndex = startIndex;
            el.textEndIndex = endIndex;

            if (prStart){
                int sp = rt.GetSpaceBefore();
                float spaceBefore;
                if(sp >= 0){
                    spaceBefore = lineHeight * sp/100;
                } else {
                    spaceBefore = -sp*Shape.POINT_DPI/Shape.MASTER_DPI;
                }
                el.ascent += spaceBefore;
            }

            float descent;
            if(linespacing >= 0){
                descent = (textLayout.GetDescent() + textLayout.GetLeading())*linespacing/100;
            } else {
                descent = -linespacing*Shape.POINT_DPI/Shape.MASTER_DPI;
            }
            if (prStart){
                int sp = rt.GetSpaceAfter();
                float spaceAfter;
                if(sp >= 0){
                    spaceAfter = lineHeight * sp/100;
                } else {
                    spaceAfter = -sp*Shape.POINT_DPI/Shape.MASTER_DPI;
                }
                el.ascent += spaceAfter;
            }
            el.descent = descent;

            if(rt.IsBullet() && (prStart || startIndex == 0)){
                it.SetIndex(startIndex);

                AttributedString bat = new AttributedString(Character.ToString(rt.GetBulletChar()));
                Color clr = rt.GetBulletColor();
                if (clr != null) bat.AddAttribute(TextAttribute.FOREGROUND, clr);
                else bat.AddAttribute(TextAttribute.FOREGROUND, it.GetAttribute(TextAttribute.FOREGROUND));

                int fontIdx = rt.GetBulletFont();
                if(fontIdx == -1) fontIdx = rt.GetFontIndex();
                PPFont bulletFont = _shape.Sheet.GetSlideShow().GetFont(fontIdx);
                bat.AddAttribute(TextAttribute.FAMILY, bulletFont.GetFontName());

                int bulletSize = rt.GetBulletSize();
                int fontSize = rt.GetFontSize();
                if(bulletSize != -1) fontSize = Math.round(fontSize*bulletSize*0.01f);
                bat.AddAttribute(TextAttribute.SIZE, new Float(fontSize));

                if(!new Font(bulletFont.GetFontName(), Font.PLAIN, 1).canDisplay(rt.GetBulletChar())){
                    bat.AddAttribute(TextAttribute.FAMILY, "Arial");
                    bat = new AttributedString("" + DEFAULT_BULLET_CHAR, bat.GetIterator().GetAttributes());
                }

                if(text.Substring(startIndex, endIndex).Length > 1){
                    el._bullet = bat;
                    el._bulletOffset = bulletOffset;
                }
            }
            lines.Add(el);
        }

        //finally Draw the text fragments
        TextElement[] elems = new TextElement[lines.Count];
        return lines.ToArray(elems);
    }

    public static class TextElement {
        public AttributedString _text;
        public int _textOffset;
        public AttributedString _bullet;
        public int _bulletOffset;
        public int _align;
        public float ascent, descent;
        public float advance;
        public int textStartIndex, textEndIndex;
    }
}





