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

namespace NPOI.XSSF.UserModel
{
    using System;
    using System.Collections.Generic;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;
    using NPOI.OpenXmlFormats.Dml.Spreadsheet;
    using NPOI.OpenXmlFormats.Dml;
    using System.Text;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.Util;


    /**
     * Represents a shape with a predefined geometry in a SpreadsheetML Drawing.
     * Possible shape types are defined in {@link NPOI.SS.UserModel.ShapeTypes}
     */
    public class XSSFSimpleShape : XSSFShape, IEnumerable<XSSFTextParagraph>
    { // TODO - instantiable superclass
        /**
         * List of the paragraphs that make up the text in this shape
         */
        private List<XSSFTextParagraph> _paragraphs;
        /**
         * A default instance of CTShape used for creating new shapes.
         */
        private static CT_Shape prototype = null;

        /**
         *  Xml bean that stores properties of this shape
         */
        private CT_Shape ctShape;

        protected internal XSSFSimpleShape(XSSFDrawing Drawing, CT_Shape ctShape)
        {
            this.drawing = Drawing;
            this.ctShape = ctShape;

            _paragraphs = new List<XSSFTextParagraph>();

            // Initialize any existing paragraphs - this will be the default body paragraph in a new shape, 
            // or existing paragraphs that have been loaded from the file
            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TextBody body = ctShape.txBody;
            if (body != null)
            {
                for (int i = 0; i < body.SizeOfPArray(); i++)
                {
                    _paragraphs.Add(new XSSFTextParagraph(body.GetPArray(i), ctShape));
                }
            }
        }

        /**
         * Prototype with the default structure of a new auto-shape.
         */
        protected internal static CT_Shape GetPrototype()
        {
            if (prototype == null)
            {
                CT_Shape shape = new CT_Shape();

                CT_ShapeNonVisual nv = shape.AddNewNvSpPr();
                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualDrawingProps nvp = nv.AddNewCNvPr();
                nvp.id = (/*setter*/1);
                nvp.name = (/*setter*/"Shape 1");
                nv.AddNewCNvSpPr();

                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties sp = shape.AddNewSpPr();
                CT_Transform2D t2d = sp.AddNewXfrm();
                CT_PositiveSize2D p1 = t2d.AddNewExt();
                p1.cx = (/*setter*/0);
                p1.cy = (/*setter*/0);
                CT_Point2D p2 = t2d.AddNewOff();
                p2.x = (/*setter*/0);
                p2.y = (/*setter*/0);

                CT_PresetGeometry2D geom = sp.AddNewPrstGeom();
                geom.prst = (/*setter*/ST_ShapeType.rect);
                geom.AddNewAvLst();

                NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TextBody body = shape.AddNewTxBody();
                CT_TextBodyProperties bodypr = body.AddNewBodyPr();
                bodypr.anchor = (/*setter*/ST_TextAnchoringType.t);
                bodypr.rtlCol = (/*setter*/false);
                CT_TextParagraph p = body.AddNewP();
                p.AddNewPPr().algn = (/*setter*/ST_TextAlignType.l);
                CT_TextCharacterProperties endPr = p.AddNewEndParaRPr();
                endPr.lang = (/*setter*/"en-US");
                endPr.sz = (/*setter*/1100);
                CT_SolidColorFillProperties scfpr = endPr.AddNewSolidFill();
                scfpr.AddNewSrgbClr().val = (/*setter*/new byte[] { 0, 0, 0 });

                body.AddNewLstStyle();

                prototype = shape;
            }
            return prototype;
        }


        public CT_Shape GetCTShape()
        {
            return ctShape;
        }


        public IEnumerator<XSSFTextParagraph> GetEnumerator()
        {
            return (IEnumerator<XSSFTextParagraph>)_paragraphs.GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
        /**
         * Returns the text from all paragraphs in the shape. Paragraphs are Separated by new lines.
         * 
         * @return  text Contained within this shape or empty string
         */
        public String Text
        {
            get
            {
                int MAX_LEVELS = 9;
                StringBuilder out1 = new StringBuilder();
                List<int> levelCount = new List<int>(MAX_LEVELS);	// maximum 9 levels
                XSSFTextParagraph p = null;

                // Initialise the levelCount array - this maintains a record of the numbering to be used at each level
                for (int k = 0; k < MAX_LEVELS; k++)
                {
                    levelCount.Add(0);
                }

                for (int i = 0; i < _paragraphs.Count; i++)
                {
                    if (out1.Length > 0) out1.Append('\n');
                    p = _paragraphs[(i)];

                    if (p.IsBullet && p.Text.Length > 0)
                    {

                        int level = Math.Min(p.Level, MAX_LEVELS - 1);

                        if (p.IsBulletAutoNumber)
                        {
                            i = ProcessAutoNumGroup(i, level, levelCount, out1);
                        }
                        else
                        {
                            // indent appropriately for the level
                            for (int j = 0; j < level; j++)
                            {
                                out1.Append('\t');
                            }
                            String character = p.BulletCharacter;
                            out1.Append(character.Length > 0 ? character + " " : "- ");
                            out1.Append(p.Text);
                        }
                    }
                    else
                    {
                        out1.Append(p.Text);

                        // this paragraph is not a bullet, so reset the count array
                        for (int k = 0; k < MAX_LEVELS; k++)
                        {
                            levelCount[k] = 0;
                        }
                    }
                }

                return out1.ToString();
            }
        }

        /**
         * 
         */
        private int ProcessAutoNumGroup(int index, int level, List<int> levelCount, StringBuilder out1)
        {
            XSSFTextParagraph p = null;
            XSSFTextParagraph nextp = null;
            ListAutoNumber scheme, nextScheme;
            int startAt, nextStartAt;

            p = _paragraphs[(index)];

            // The rules for generating the auto numbers are as follows. If the following paragraph is also
            // an auto-number, has the same type/scheme (and startAt if defined on this paragraph) then they are
            // considered part of the same group. An empty bullet paragraph is counted as part of the same
            // group but does not increment the count for the group. A change of type, startAt or the paragraph
            // not being a bullet resets the count for that level to 1.

            // first auto-number paragraph so Initialise to 1 or the bullets startAt if present
            startAt = p.BulletAutoNumberStart;
            scheme = p.BulletAutoNumberScheme;
            if (levelCount[(level)] == 0)
            {
                levelCount[level] = startAt == 0 ? 1 : startAt;
            }
            // indent appropriately for the level
            for (int j = 0; j < level; j++)
            {
                out1.Append('\t');
            }
            if (p.Text.Length > 0)
            {
                out1.Append(GetBulletPrefix(scheme, levelCount[level]));
                out1.Append(p.Text);
            }
            while (true)
            {
                nextp = (index + 1) == _paragraphs.Count ? null : _paragraphs[(index + 1)];
                if (nextp == null) break; // out of paragraphs
                if (!(nextp.IsBullet && p.IsBulletAutoNumber)) break; // not an auto-number bullet                      
                if (nextp.Level > level)
                {
                    // recurse into the new level group
                    if (out1.Length > 0) out1.Append('\n');
                    index = ProcessAutoNumGroup(index + 1, nextp.Level, levelCount, out1);
                    continue; // restart the loop given the new index
                }
                else if (nextp.Level < level)
                {
                    break; // Changed level   
                }
                nextScheme = nextp.BulletAutoNumberScheme;
                nextStartAt = nextp.BulletAutoNumberStart;

                if (nextScheme == scheme && nextStartAt == startAt)
                {
                    // bullet is valid, so increment i 
                    ++index;
                    if (out1.Length > 0) out1.Append('\n');
                    // indent for the level
                    for (int j = 0; j < level; j++)
                    {
                        out1.Append('\t');
                    }
                    // check for empty text - only output a bullet if there is text, but it is still part of the group
                    if (nextp.Text.Length > 0)
                    {
                        // increment the count for this level
                        levelCount[level] = levelCount[(level) + 1];
                        out1.Append(GetBulletPrefix(nextScheme, levelCount[(level)]));
                        out1.Append(nextp.Text);
                    }
                }
                else
                {
                    // something doesn't match so stop
                    break;
                }
            }
            // end of the group so reset the count for this level 
            levelCount[level] = 0;

            return index;
        }
        /**
         * Returns a string Containing an appropriate prefix for an auto-numbering bullet
         * @param scheme the auto-numbering scheme used by the bullet
         * @param value the value of the bullet
         * @return appropriate prefix for an auto-numbering bullet
         */
        private String GetBulletPrefix(ListAutoNumber scheme, int value)
        {
            StringBuilder out1 = new StringBuilder();

            switch (scheme)
            {
                case ListAutoNumber.ALPHA_LC_PARENT_BOTH:
                case ListAutoNumber.ALPHA_LC_PARENT_R:
                    if (scheme == ListAutoNumber.ALPHA_LC_PARENT_BOTH) out1.Append('(');
                    out1.Append(valueToAlpha(value).ToLower());
                    out1.Append(')');
                    break;
                case ListAutoNumber.ALPHA_UC_PARENT_BOTH:
                case ListAutoNumber.ALPHA_UC_PARENT_R:
                    if (scheme == ListAutoNumber.ALPHA_UC_PARENT_BOTH) out1.Append('(');
                    out1.Append(valueToAlpha(value));
                    out1.Append(')');
                    break;
                case ListAutoNumber.ALPHA_LC_PERIOD:
                    out1.Append(valueToAlpha(value).ToLower());
                    out1.Append('.');
                    break;
                case ListAutoNumber.ALPHA_UC_PERIOD:
                    out1.Append(valueToAlpha(value));
                    out1.Append('.');
                    break;
                case ListAutoNumber.ARABIC_PARENT_BOTH:
                case ListAutoNumber.ARABIC_PARENT_R:
                    if (scheme == ListAutoNumber.ARABIC_PARENT_BOTH) out1.Append('(');
                    out1.Append(value);
                    out1.Append(')');
                    break;
                case ListAutoNumber.ARABIC_PERIOD:
                    out1.Append(value);
                    out1.Append('.');
                    break;
                case ListAutoNumber.ARABIC_PLAIN:
                    out1.Append(value);
                    break;
                case ListAutoNumber.ROMAN_LC_PARENT_BOTH:
                case ListAutoNumber.ROMAN_LC_PARENT_R:
                    if (scheme == ListAutoNumber.ROMAN_LC_PARENT_BOTH) out1.Append('(');
                    out1.Append(valueToRoman(value).ToLower());
                    out1.Append(')');
                    break;
                case ListAutoNumber.ROMAN_UC_PARENT_BOTH:
                case ListAutoNumber.ROMAN_UC_PARENT_R:
                    if (scheme == ListAutoNumber.ROMAN_UC_PARENT_BOTH) out1.Append('(');
                    out1.Append(valueToRoman(value));
                    out1.Append(')');
                    break;
                case ListAutoNumber.ROMAN_LC_PERIOD:
                    out1.Append(valueToRoman(value).ToLower());
                    out1.Append('.');
                    break;
                case ListAutoNumber.ROMAN_UC_PERIOD:
                    out1.Append(valueToRoman(value));
                    out1.Append('.');
                    break;
                default:
                    out1.Append('\u2022');   // can't Set the font to wingdings so use the default bullet character
                    break;
            }
            out1.Append(" ");
            return out1.ToString();
        }

        /**
         * Convert an integer to its alpha equivalent e.g. 1 = A, 2 = B, 27 = AA etc
         */
        private String valueToAlpha(int value)
        {
            String alpha = "";
            int modulo;
            while (value > 0)
            {
                modulo = (value - 1) % 26;
                alpha = (char)(65 + modulo) + alpha;
                value = (value - modulo) / 26;
            }
            return alpha;
        }

        private static String[] _romanChars = new String[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        private static int[] _romanAlphaValues = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };

        /**
         * Convert an integer to its roman equivalent e.g. 1 = I, 9 = IX etc
         */
        private String valueToRoman(int value)
        {
            StringBuilder out1 = new StringBuilder();
            for (int i = 0; value > 0 && i < _romanChars.Length; i++)
            {
                while (_romanAlphaValues[i] <= value)
                {
                    out1.Append(_romanChars[i]);
                    value -= _romanAlphaValues[i];
                }
            }
            return out1.ToString();
        }

        /**
         * Clear all text from this shape
         */
        public void ClearText()
        {
            _paragraphs.Clear();
            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TextBody txBody = ctShape.txBody;
            txBody.SetPArray(null); // remove any existing paragraphs
        }

        /**
         * Set a single paragraph of text on the shape. Note this will replace all existing paragraphs Created on the shape.
         * @param text	string representing the paragraph text
         */
        public void SetText(String text)
        {
            ClearText();

            AddNewTextParagraph().AddNewTextRun().Text = (text);
        }

        /**
         * Set a single paragraph of text on the shape. Note this will replace all existing paragraphs Created on the shape.
         * @param str	rich text string representing the paragraph text
         */
        public void SetText(XSSFRichTextString str)
        {

            XSSFWorkbook wb = (XSSFWorkbook)GetDrawing().GetParent().GetParent();
            str.SetStylesTableReference(wb.GetStylesSource());

            CT_TextParagraph p = new CT_TextParagraph();
            if (str.NumFormattingRuns == 0)
            {
                CT_RegularTextRun r = p.AddNewR();
                CT_TextCharacterProperties rPr = r.AddNewRPr();
                rPr.lang = (/*setter*/"en-US");
                rPr.sz = (/*setter*/1100);
                r.t = (/*setter*/str.String);

            }
            else
            {
                for (int i = 0; i < str.GetCTRst().SizeOfRArray(); i++)
                {
                    CT_RElt lt = str.GetCTRst().GetRArray(i);
                    CT_RPrElt ltPr = lt.rPr;
                    if (ltPr == null) ltPr = lt.AddNewRPr();

                    CT_RegularTextRun r = p.AddNewR();
                    CT_TextCharacterProperties rPr = r.AddNewRPr();
                    rPr.lang = (/*setter*/"en-US");

                    ApplyAttributes(ltPr, rPr);

                    r.t = (/*setter*/lt.t);
                }
            }

            ClearText();
            ctShape.txBody.SetPArray(new CT_TextParagraph[] { p });
            _paragraphs.Add(new XSSFTextParagraph(ctShape.txBody.GetPArray(0), ctShape));
        }

        /**
         * Returns a collection of the XSSFTextParagraphs that are attached to this shape
         * 
         * @return text paragraphs in this shape
         */
        public List<XSSFTextParagraph> TextParagraphs
        {
            get
            {
                return _paragraphs;
            }
        }

        /**
         * Add a new paragraph run to this shape
         *
         * @return Created paragraph run
         */
        public XSSFTextParagraph AddNewTextParagraph()
        {
            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TextBody txBody = ctShape.txBody;
            CT_TextParagraph p = txBody.AddNewP();
            XSSFTextParagraph paragraph = new XSSFTextParagraph(p, ctShape);
            _paragraphs.Add(paragraph);
            return paragraph;
        }

        /**
         * Add a new paragraph run to this shape, Set to the provided string
         *
         * @return Created paragraph run
         */
        public XSSFTextParagraph AddNewTextParagraph(String text)
        {
            XSSFTextParagraph paragraph = AddNewTextParagraph();
            paragraph.AddNewTextRun().Text=(text);
            return paragraph;
        }

        /**
         * Add a new paragraph run to this shape, Set to the provided rich text string 
         *
         * @return Created paragraph run
         */
        public XSSFTextParagraph AddNewTextParagraph(XSSFRichTextString str)
        {
            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_TextBody txBody = ctShape.txBody;
            CT_TextParagraph p = txBody.AddNewP();

            if (str.NumFormattingRuns == 0)
            {
                CT_RegularTextRun r = p.AddNewR();
                CT_TextCharacterProperties rPr = r.AddNewRPr();
                rPr.lang = (/*setter*/"en-US");
                rPr.sz = (/*setter*/1100);
                r.t = (/*setter*/str.String);

            }
            else
            {
                for (int i = 0; i < str.GetCTRst().SizeOfRArray(); i++)
                {
                    CT_RElt lt = str.GetCTRst().GetRArray(i);
                    CT_RPrElt ltPr = lt.rPr;
                    if (ltPr == null) ltPr = lt.AddNewRPr();

                    CT_RegularTextRun r = p.AddNewR();
                    CT_TextCharacterProperties rPr = r.AddNewRPr();
                    rPr.lang = (/*setter*/"en-US");

                    ApplyAttributes(ltPr, rPr);

                    r.t = (/*setter*/lt.t);
                }
            }

            // Note: the XSSFTextParagraph constructor will create its required XSSFTextRuns from the provided CTTextParagraph
            XSSFTextParagraph paragraph = new XSSFTextParagraph(p, ctShape);
            _paragraphs.Add(paragraph);

            return paragraph;
        }

        /**
         * Returns the type of horizontal overflow for the text.
         *
         * @return the type of horizontal overflow
         */
        public TextHorizontalOverflow TextHorizontalOverflow
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetHorzOverflow())
                    {
                        return (TextHorizontalOverflow)((int)bodyPr.horzOverflow - 1);
                    }
                }
                return TextHorizontalOverflow.OVERFLOW;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (value == TextHorizontalOverflow.None)
                    {
                        if (bodyPr.IsSetHorzOverflow()) bodyPr.UnsetHorzOverflow();
                    }
                    else
                    {
                        bodyPr.horzOverflow = (/*setter*/(ST_TextHorzOverflowType)((int)value + 1));
                    }
                }
            }
        }

        /**
         * Returns the type of vertical overflow for the text.
         *
         * @return the type of vertical overflow
         */
        public TextVerticalOverflow TextVerticalOverflow
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetVertOverflow())
                    {
                        return (TextVerticalOverflow)((int)bodyPr.vertOverflow - 1);
                    }
                }
                return TextVerticalOverflow.OVERFLOW;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (value == TextVerticalOverflow.None)
                    {
                        if (bodyPr.IsSetVertOverflow()) bodyPr.UnsetVertOverflow();
                    }
                    else
                    {
                        bodyPr.vertOverflow = (/*setter*/(ST_TextVertOverflowType)((int)value + 1));
                    }
                }
            }
        }

        /**
         * Returns the type of vertical alignment for the text within the shape.
         *
         * @return the type of vertical alignment
         */
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetAnchor())
                    {
                        return (VerticalAlignment)((int)bodyPr.anchor);
                    }
                }
                return VerticalAlignment.Top;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (value == VerticalAlignment.None)
                    {
                        if (bodyPr.IsSetAnchor()) bodyPr.UnsetAnchor();
                    }
                    else
                    {
                        bodyPr.anchor = (/*setter*/(ST_TextAnchoringType)((int)value));
                    }
                }
            }
        }

        /**
         * Gets the vertical orientation of the text
         * 
         * @return vertical orientation of the text
         */
        public TextDirection TextDirection
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    ST_TextVerticalType val = bodyPr.vert;
                    if (val != ST_TextVerticalType.horz)
                    {
                        return (TextDirection)(val - 1);
                    }
                }
                return TextDirection.HORIZONTAL;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (value == TextDirection.None)
                    {
                        if (bodyPr.IsSetVert()) bodyPr.UnsetVert();
                    }
                    else
                    {
                        bodyPr.vert = (/*setter*/(ST_TextVerticalType)((int)value + 1));
                    }
                }
            }
        }


        /**
         * Returns the distance (in points) between the bottom of the text frame
         * and the bottom of the inscribed rectangle of the shape that Contains the text.
         *
         * @return the bottom inset in points
         */
        public double BottomInset
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetBIns())
                    {
                        return Units.ToPoints(bodyPr.bIns);
                    }
                }
                // If this attribute is omitted, then a value of 0.05 inches is implied
                return 3.6;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (value == -1)
                    {
                        if (bodyPr.IsSetBIns()) bodyPr.UnsetBIns();
                    }
                    else bodyPr.bIns = (/*setter*/Units.ToEMU(value));
                }
            }
        }

        /**
         *  Returns the distance (in points) between the left edge of the text frame
         *  and the left edge of the inscribed rectangle of the shape that Contains
         *  the text.
         *
         * @return the left inset in points
         */
        public double LeftInset
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetLIns())
                    {
                        return Units.ToPoints(bodyPr.lIns);
                    }
                }
                // If this attribute is omitted, then a value of 0.05 inches is implied
                return 3.6;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (value == -1)
                    {
                        if (bodyPr.IsSetLIns()) bodyPr.UnsetLIns();
                    }
                    else bodyPr.lIns = (/*setter*/Units.ToEMU(value));
                }
            }
        }

        /**
         *  Returns the distance (in points) between the right edge of the
         *  text frame and the right edge of the inscribed rectangle of the shape
         *  that Contains the text.
         *
         * @return the right inset in points
         */
        public double RightInset
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetRIns())
                    {
                        return Units.ToPoints(bodyPr.rIns);
                    }
                }
                // If this attribute is omitted, then a value of 0.05 inches is implied
                return 3.6;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (value == -1)
                    {
                        if (bodyPr.IsSetRIns()) bodyPr.UnsetRIns();
                    }
                    else bodyPr.rIns = (/*setter*/Units.ToEMU(value));
                }
            }
        }

        /**
         *  Returns the distance (in points) between the top of the text frame
         *  and the top of the inscribed rectangle of the shape that Contains the text.
         *
         * @return the top inset in points
         */
        public double TopInset
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetTIns())
                    {
                        return Units.ToPoints(bodyPr.tIns);
                    }
                }
                // If this attribute is omitted, then a value of 0.05 inches is implied
                return 3.6;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (value == -1)
                    {
                        if (bodyPr.IsSetTIns()) bodyPr.UnsetTIns();
                    }
                    else bodyPr.tIns = (/*setter*/Units.ToEMU(value));
                }
            }
        }

        /**
         * @return whether to wrap words within the bounding rectangle
         */
        public bool WordWrap
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetWrap())
                    {
                        return bodyPr.wrap == ST_TextWrappingType.square;
                    }
                }
                return true;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    bodyPr.wrap = (/*setter*/value ? ST_TextWrappingType.square : ST_TextWrappingType.none);
                }
            }
        }



        /**
         *
         * Specifies that a shape should be auto-fit to fully contain the text described within it.
         * Auto-fitting is when text within a shape is scaled in order to contain all the text inside
         *
         * @param value type of autofit
         * @return type of autofit
         */
        public TextAutofit TextAutofit
        {
            get
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetNoAutofit()) return TextAutofit.NONE;
                    else if (bodyPr.IsSetNormAutofit()) return TextAutofit.NORMAL;
                    else if (bodyPr.IsSetSpAutoFit()) return TextAutofit.SHAPE;
                }
                return TextAutofit.NORMAL;
            }
            set
            {
                CT_TextBodyProperties bodyPr = ctShape.txBody.bodyPr;
                if (bodyPr != null)
                {
                    if (bodyPr.IsSetSpAutoFit()) bodyPr.UnsetSpAutoFit();
                    if (bodyPr.IsSetNoAutofit()) bodyPr.UnsetNoAutofit();
                    if (bodyPr.IsSetNormAutofit()) bodyPr.UnsetNormAutofit();

                    switch (value)
                    {
                        case TextAutofit.NONE: bodyPr.AddNewNoAutofit(); break;
                        case TextAutofit.NORMAL: bodyPr.AddNewNormAutofit(); break;
                        case TextAutofit.SHAPE: bodyPr.AddNewSpAutoFit(); break;
                    }
                }
            }
        }

        /**
         * Gets the shape type, one of the constants defined in {@link NPOI.SS.UserModel.ShapeTypes}.
         *
         * @return the shape type
         * @see NPOI.SS.UserModel.ShapeTypes
         */
        public int ShapeType
        {
            get
            {
                return (int)ctShape.spPr.prstGeom.prst;
            }
            set
            {
                ctShape.spPr.prstGeom.prst = (/*setter*/(ST_ShapeType)(value));
            }
        }

        protected internal override NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties GetShapeProperties()
        {
            return ctShape.spPr;
        }

        /**
         * org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTRPrElt to
         * org.Openxmlformats.schemas.Drawingml.x2006.main.CTFont adapter
         */
        private static void ApplyAttributes(CT_RPrElt pr, CT_TextCharacterProperties rPr)
        {

            if (pr.SizeOfBArray() > 0) rPr.b = (/*setter*/pr.GetBArray(0).val);
            if (pr.SizeOfUArray() > 0)
            {
                ST_UnderlineValues u1 = pr.GetUArray(0).val;
                if (u1 == ST_UnderlineValues.single) rPr.u = (/*setter*/ST_TextUnderlineType.sng);
                else if (u1 == ST_UnderlineValues.@double) rPr.u = (/*setter*/ST_TextUnderlineType.dbl);
                else if (u1 == ST_UnderlineValues.none) rPr.u = (/*setter*/ST_TextUnderlineType.none);
            }
            if (pr.SizeOfIArray() > 0) rPr.i = (/*setter*/pr.GetIArray(0).val);

            if (pr.SizeOfRFontArray() > 0)
            {
                CT_TextFont rFont = rPr.IsSetLatin() ? rPr.latin : rPr.AddNewLatin();
                rFont.typeface = (/*setter*/pr.GetRFontArray(0).val);
            }

            if (pr.SizeOfSzArray() > 0)
            {
                int sz = (int)(pr.GetSzArray(0).val * 100);
                rPr.sz = (/*setter*/sz);
            }

            if (pr.SizeOfColorArray() > 0)
            {
                CT_SolidColorFillProperties fill = rPr.IsSetSolidFill() ? rPr.solidFill : rPr.AddNewSolidFill();
                NPOI.OpenXmlFormats.Spreadsheet.CT_Color xlsColor = pr.GetColorArray(0);
                if (xlsColor.IsSetRgb())
                {
                    CT_SRgbColor clr = fill.IsSetSrgbClr() ? fill.srgbClr : fill.AddNewSrgbClr();
                    clr.val = (/*setter*/xlsColor.rgb);
                }
                else if (xlsColor.IsSetIndexed())
                {
                    HSSFColor indexed = (HSSFColor)HSSFColor.GetIndexHash()[((int)xlsColor.indexed)];
                    if (indexed != null)
                    {
                        byte[] rgb = new byte[3];
                        rgb[0] = (byte)indexed.GetTriplet()[0];
                        rgb[1] = (byte)indexed.GetTriplet()[1];
                        rgb[2] = (byte)indexed.GetTriplet()[2];
                        CT_SRgbColor clr = fill.IsSetSrgbClr() ? fill.srgbClr : fill.AddNewSrgbClr();
                        clr.val = (/*setter*/rgb);
                    }
                }
            }
        }
    }
}