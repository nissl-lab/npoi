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
namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;
    using System.Text;
    using System.Xml;
    using System.IO;
    using NPOI.Util;
    using NPOI.OpenXmlFormats.Dml;
    using System.Xml.Serialization;
    using NPOI.OpenXmlFormats.Dml.WordProcessing;
    using NPOI.WP.UserModel;

    /**
     * @see <a href="http://msdn.microsoft.com/en-us/library/ff533743(v=office.12).aspx">[MS-OI29500] Run Fonts</a> 
     */
    public enum FontCharRange
    {
        None,
        Ascii /* char 0-127 */,
        CS /* complex symbol */,
        EastAsia /* east asia */,
        HAnsi /* high ansi */
    };
    /**
     * XWPFrun.object defines a region of text with a common Set of properties
     *
     * @author Yegor Kozlov
     * @author Gregg Morris (gregg dot morris at gmail dot com) - added getColor(), setColor()
     */
    public class XWPFRun : ISDTContents, IRunElement, ICharacterRun
    {
        private CT_R run;
        private String pictureText;
        //private XWPFParagraph paragraph;
        private IRunBody parent;
        private List<XWPFPicture> pictures;

        /**
         * @param r the CT_R bean which holds the run.attributes
         * @param p the parent paragraph
         */
        public XWPFRun(CT_R r, IRunBody p)
        {
            this.run = r;
            this.parent = p;

            /**
             * reserve already occupied Drawing ids, so reserving new ids later will
             * not corrupt the document
             */
            IList<CT_Drawing> drawingList = r.GetDrawingList();
            foreach (CT_Drawing ctDrawing in drawingList)
            {
                List<CT_Anchor> anchorList = ctDrawing.GetAnchorList();
                foreach (CT_Anchor anchor in anchorList)
                {
                    if (anchor.docPr != null)
                    {
                        this.Document.DrawingIdManager.Reserve(anchor.docPr.id);
                    }
                }
                List<CT_Inline> inlineList = ctDrawing.GetInlineList();
                foreach (CT_Inline inline in inlineList)
                {
                    if (inline.docPr != null)
                    {
                        this.Document.DrawingIdManager.Reserve(inline.docPr.id);
                    }
                }
            }

            //// Look for any text in any of our pictures or Drawings
            StringBuilder text = new StringBuilder();
            List<object> pictTextObjs = new List<object>();
            foreach (CT_Picture pic in r.GetPictList())
                pictTextObjs.Add(pic);
            foreach (CT_Drawing draw in drawingList)
                pictTextObjs.Add(draw);
            //foreach (object o in pictTextObjs)
            //{
            //todo:: imlement this
            //XmlObject[] t = o.SelectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:t");
            //for (int m = 0; m < t.Length; m++)
            //{
            //    NodeList kids = t[m].DomNode.ChildNodes;
            //    for (int n = 0; n < kids.Length; n++)
            //    {
            //        if (kids.Item(n) is Text)
            //        {
            //            if (text.Length > 0)
            //                text.Append("\n");
            //            text.Append(kids.Item(n).NodeValue);
            //        }
            //    }
            //}
            //}
            pictureText = text.ToString();

            // Do we have any embedded pictures?
            // (They're a different CT_Picture, under the Drawingml namespace)
            pictures = new List<XWPFPicture>();
            foreach (object o in pictTextObjs)
            {
                foreach (OpenXmlFormats.Dml.Picture.CT_Picture pict in GetCTPictures(o))
                {
                    XWPFPicture picture = new XWPFPicture(pict, this);
                    pictures.Add(picture);
                }
            }
        }

        /**
         * @deprecated Use {@link XWPFRun#XWPFRun(CTR, IRunBody)}
         */
        [Obsolete("Use XWPFRun(CTR, IRunBody)")]
        public XWPFRun(CT_R r, XWPFParagraph p)
            : this(r, (IRunBody)p)
        {
        }
        private List<NPOI.OpenXmlFormats.Dml.Picture.CT_Picture> GetCTPictures(object o)
        {
            List<NPOI.OpenXmlFormats.Dml.Picture.CT_Picture> pictures = new List<NPOI.OpenXmlFormats.Dml.Picture.CT_Picture>();
            //XmlObject[] picts = o.SelectPath("declare namespace pic='"+CT_Picture.type.Name.NamespaceURI+"' .//pic:pic");
            //XmlElement[] picts = o.Any;
            //foreach (XmlElement pict in picts)
            //{
            //if(pict is XmlAnyTypeImpl) {
            //    // Pesky XmlBeans bug - see Bugzilla #49934
            //    try {
            //        pict = CT_Picture.Factory.Parse( pict.ToString() );
            //    } catch(XmlException e) {
            //        throw new POIXMLException(e);
            //    }
            //}
            //if (pict is NPOI.OpenXmlFormats.Dml.CT_Picture)
            //{
            //    pictures.Add((NPOI.OpenXmlFormats.Dml.CT_Picture)pict);
            //}
            //}
            if (o is CT_Drawing)
            {
                CT_Drawing drawing = o as CT_Drawing;
                if (drawing.inline != null)
                {
                    foreach (CT_Inline inline in drawing.inline)
                    {
                        GetPictures(inline.graphic.graphicData, pictures);
                    }
                }
            }
            else if (o is CT_GraphicalObjectData)
            {
                GetPictures(o as CT_GraphicalObjectData, pictures);
            }
            return pictures;
        }

        private void GetPictures(CT_GraphicalObjectData god, List<NPOI.OpenXmlFormats.Dml.Picture.CT_Picture> pictures)
        {
            XmlSerializer xmlse = new XmlSerializer(typeof(NPOI.OpenXmlFormats.Dml.Picture.CT_Picture));
            foreach (string el in god.Any)
            {
                System.IO.StringReader stringReader = new System.IO.StringReader(el);

                NPOI.OpenXmlFormats.Dml.Picture.CT_Picture pict =
                    xmlse.Deserialize(System.Xml.XmlReader.Create(stringReader)) as NPOI.OpenXmlFormats.Dml.Picture.CT_Picture;
                pictures.Add(pict);
            }
        }

        /**
         * Get the currently used CT_R object
         * @return CT_R object
         */

        public CT_R GetCTR()
        {
            return run;
        }

        /**
         * Get the currently referenced paragraph/SDT object
         * @return current parent
         */
        public IRunBody Parent
        {
            get
            {
                return parent;
            }
        }
        /**
         * Get the currently referenced paragraph, or null if a SDT object
         * @deprecated use {@link XWPFRun#getParent()} instead
         */
        public XWPFParagraph Paragraph
        {
            get
            {
                if (parent is XWPFParagraph)
                    return (XWPFParagraph)parent;
                return null;
            }
        }

        /**
         * @return The {@link XWPFDocument} instance, this run.belongs to, or
         *         <code>null</code> if parent structure (paragraph > document) is not properly Set.
         */
        public XWPFDocument Document
        {
            get
            {
                if (parent != null)
                {
                    return parent.Document;
                }
                return null;
            }
        }

        /**
         * For isBold, isItalic etc
         */
        private bool IsCTOnOff(CT_OnOff onoff)
        {
            if (!onoff.IsSetVal())
                return true;
            return onoff.val;
        }

        /**
         * Whether the bold property shall be applied to all non-complex script
         * characters in the contents of this run.when displayed in a document. 
         * <p>
         * This formatting property is a toggle property, which specifies that its
         * behavior differs between its use within a style defInition and its use as
         * direct formatting. When used as part of a style defInition, Setting this
         * property shall toggle the current state of that property as specified up
         * to this point in the hierarchy (i.e. applied to not applied, and vice
         * versa). Setting it to <code>false</code> (or an equivalent) shall
         * result in the current Setting remaining unChanged. However, when used as
         * direct formatting, Setting this property to true or false shall Set the
         * absolute state of the resulting property.
         * </p>
         * <p>
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then bold shall not be
         * applied to non-complex script characters.
         * </p>
         *
         * @param value <code>true</code> if the bold property is applied to
         *              this run
         */
        public bool IsBold
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetB())
                {
                    return false;
                }
                return IsCTOnOff(pr.b);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff bold = pr.IsSetB() ? pr.b : pr.AddNewB();
                bold.val = value;
            }
        }

        /**
     * Get text color. The returned value is a string in the hex form "RRGGBB".
     */
        public String GetColor()
        {
            String color = null;
            if (run.IsSetRPr())
            {
                CT_RPr pr = run.rPr;
                if (pr.IsSetColor())
                {
                    NPOI.OpenXmlFormats.Wordprocessing.CT_Color clr = pr.color;
                    color = clr.val; //clr.xgetVal().getStringValue();
                }
            }
            return color;
        }

        /**
         * Set text color.
         * @param rgbStr - the desired color, in the hex form "RRGGBB".
         */
        public void SetColor(String rgbStr)
        {
            CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
            NPOI.OpenXmlFormats.Wordprocessing.CT_Color color = pr.IsSetColor() ? pr.color : pr.AddNewColor();
            color.val = (rgbStr);
        }
        /**
         * Return the string content of this text run
         *
         * @return the text of this text run.or <code>null</code> if not Set
         */
        public String GetText(int pos)
        {
            return run.SizeOfTArray() == 0 ? null : run.GetTArray(pos).Value;
        }

        /**
         * Returns text embedded in pictures
         */
        public String PictureText
        {
            get
            {
                return pictureText;
            }
        }
        public void ReplaceText(string oldText, string newText)
        {
            string text = this.Text.Replace(oldText, newText);
            this.SetText(text);
        }
        /// <summary>
        ///Sets the text of this text run
        /// </summary>
        /// <param name="value">the literal text which shall be displayed in the document</param>
        public void SetText(String value)
        {
            SetText(value, 0);
        }


        public void AppendText(String value)
        {
            SetText(value, run.GetTList().Count);
        }

        /**
         * Sets the text of this text run.in the 
         *
         * @param value the literal text which shall be displayed in the document
         * @param pos - position in the text array (NB: 0 based)
         */
        public void SetText(String value, int pos)
        {
            int length = run.SizeOfTArray();
            if (pos > length) throw new IndexOutOfRangeException("Value too large for the parameter position");
            CT_Text t = (pos < length && pos >= 0) ? run.GetTArray(pos): run.AddNewT();
            t.Value = (value);
            preserveSpaces(t);
        }

        /**
         * Whether the italic property should be applied to all non-complex script
         * characters in the contents of this run.when displayed in a document.
         *
         * @return <code>true</code> if the italic property is applied
         */
        public bool IsItalic
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetI())
                    return false;
                return IsCTOnOff(pr.i);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff italic = pr.IsSetI() ? pr.i : pr.AddNewI();
                italic.val = value;
            }
        }


        /**
         * Specifies that the contents of this run.should be displayed along with an
         * underline appearing directly below the character heigh
         *
         * @return the Underline pattern Applyed to this run
         * @see UnderlinePatterns
         */
        public UnderlinePatterns Underline
        {
            get
            {
                CT_RPr pr = run.rPr;
                return (pr != null && pr.IsSetU() && pr.u.val != null) ?
                    EnumConverter.ValueOf<UnderlinePatterns, ST_Underline>(pr.u.val) : UnderlinePatterns.None;
            }
        }
        internal void InsertText(CT_Text text, int textIndex)
        {
            run.GetTList().Insert(textIndex, text);
        }

        /// <summary>
        /// insert text at start index in the run
        /// </summary>
        /// <param name="text">insert text</param>
        /// <param name="startIndex">start index of the insertion in the run text</param>
        public void InsertText(string text, int startIndex)
        {
            List<CT_Text> texts = run.GetTList();
            int endPos = 0;
            int startPos = 0;
            for (int i = 0; i < texts.Count; i++)
            {
                startPos = endPos;
                endPos += texts[i].Value.Length;
                if (endPos > startIndex)
                {
                    texts[i].Value = texts[i].Value.Insert(startIndex - startPos, text);
                    break;
                }
            }
        }
        public string Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                for (int i = 0; i < run.Items.Count; i++)
                {
                    object o = run.Items[i];
                    if (o is CT_Text)
                    {
                        if (!(run.ItemsElementName[i] == RunItemsChoiceType.instrText))
                        {
                            text.Append(((CT_Text)o).Value);
                        }
                    }
                    // Complex type evaluation (currently only for extraction of check boxes)
                    if (o is CT_FldChar)
                    {
                        CT_FldChar ctfldChar = ((CT_FldChar)o);
                        if (ctfldChar.fldCharType == ST_FldCharType.begin)
                        {
                            if (ctfldChar.ffData != null)
                            {
                                foreach (CT_FFCheckBox checkBox in ctfldChar.ffData.GetCheckBoxList())
                                {
                                    if (checkBox.@default.val == true)
                                    {
                                        text.Append("|X|");
                                    }
                                    else
                                    {
                                        text.Append("|_|");
                                    }
                                }
                            }
                        }
                    }
                    if (o is CT_PTab)
                    {
                        text.Append("\t");
                    }
                    if (o is CT_Br)
                    {
                        text.Append("\n");
                    }
                    if (o is CT_Empty)
                    {
                        // Some inline text elements Get returned not as
                        //  themselves, but as CTEmpty, owing to some odd
                        //  defInitions around line 5642 of the XSDs
                        // This bit works around it, and replicates the above
                        //  rules for that case
                        if (run.ItemsElementName[i] == RunItemsChoiceType.tab)
                        {
                            text.Append("\t");
                        }
                        if (run.ItemsElementName[i] == RunItemsChoiceType.br)
                        {
                            text.Append("\n");
                        }
                        if (run.ItemsElementName[i] == RunItemsChoiceType.cr)
                        {
                            text.Append("\n");
                        }
                    }

                    if (o is CT_FtnEdnRef)
                    {
                        CT_FtnEdnRef ftn = (CT_FtnEdnRef)o;
                        String footnoteRef = ftn.DomNode.LocalName.Equals("footnoteReference") ?
                            "[footnoteRef:" + ftn.id + "]" : "[endnoteRef:" + ftn.id + "]";
                        text.Append(footnoteRef);
                    }
                }

                // Any picture text?
                if (pictureText != null && pictureText.Length > 0)
                {
                    text.Append("\n").Append(pictureText);
                }

                return text.ToString();
            }
        }
        /**
         * Specifies that the contents of this run.should be displayed along with an
         * underline appearing directly below the character heigh
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then an underline shall
         * not be applied to the contents of this run.
         *
         * @param value -
         *              underline type
         * @see UnderlinePatterns : all possible patterns that could be applied
         */
        public void SetUnderline(UnderlinePatterns value)
        {
            CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
            CT_Underline underline = (pr.u == null) ? pr.AddNewU() : pr.u;
            underline.val = EnumConverter.ValueOf<ST_Underline, UnderlinePatterns>(value);
        }

        /**
         * Specifies that the contents of this run.shall be displayed with a single
         * horizontal line through the center of the line.
         *
         * @return <code>true</code> if the strike property is applied
         */
        public bool IsStrikeThrough
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetStrike())
                    return false;
                return IsCTOnOff(pr.strike);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff strike = pr.IsSetStrike() ? pr.strike : pr.AddNewStrike();
                strike.val = value;//(value ? ST_OnOff.True : ST_OnOff.False);
            }
        }
        /**
         * Specifies that the contents of this run.shall be displayed with a single
         * horizontal line through the center of the line.
         * <p/>
         * This formatting property is a toggle property, which specifies that its
         * behavior differs between its use within a style defInition and its use as
         * direct formatting. When used as part of a style defInition, Setting this
         * property shall toggle the current state of that property as specified up
         * to this point in the hierarchy (i.e. applied to not applied, and vice
         * versa). Setting it to false (or an equivalent) shall result in the
         * current Setting remaining unChanged. However, when used as direct
         * formatting, Setting this property to true or false shall Set the absolute
         * state of the resulting property.
         * </p>
         * <p/>
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then strikethrough shall
         * not be applied to the contents of this run.
         * </p>
         *
         * @param value <code>true</code> if the strike property is applied to
         *              this run
         */
        [Obsolete]
        public bool IsStrike
        {
            get
            {
                return IsStrikeThrough;
            }
            set
            {
                IsStrikeThrough = value;
            }
        }
        /**
         * Specifies that the contents of this run shall be displayed with a double
         * horizontal line through the center of the line.
         *
         * @return <code>true</code> if the double strike property is applied
         */
        public bool IsDoubleStrikeThrough
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetDstrike())
                    return false;
                return IsCTOnOff(pr.dstrike);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff dstrike = pr.IsSetDstrike() ? pr.dstrike : pr.AddNewDstrike();
                dstrike.val = value;//(value ? STOnOff.TRUE : STOnOff.FALSE);
            }
        }
        public bool IsSmallCaps
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetSmallCaps())
                    return false;
                return IsCTOnOff(pr.smallCaps);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff caps = pr.IsSetSmallCaps() ? pr.smallCaps : pr.AddNewSmallCaps();
                caps.val = value;//(value ? ST_OnOff.True : ST_OnOff.False);
            }
        }

        public bool IsCapitalized
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetCaps())
                    return false;
                return IsCTOnOff(pr.caps);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff caps = pr.IsSetCaps() ? pr.caps : pr.AddNewCaps();
                caps.val = value;//(value ? ST_OnOff.True : ST_OnOff.False);
            }
        }

        public bool IsShadowed
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetShadow())
                    return false;
                return IsCTOnOff(pr.shadow);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff shadow = pr.IsSetShadow() ? pr.shadow : pr.AddNewShadow();
                shadow.val = value;//(value ? ST_OnOff.True : ST_OnOff.False);
            }
        }

        public bool IsImprinted
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetImprint())
                    return false;
                return IsCTOnOff(pr.imprint);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff imprinted = pr.IsSetImprint() ? pr.imprint : pr.AddNewImprint();
                imprinted.val = value;//(value ? ST_OnOff.True : ST_OnOff.False);
            }
        }

        public bool IsEmbossed
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetEmboss())
                    return false;
                return IsCTOnOff(pr.emboss);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_OnOff emboss = pr.IsSetEmboss() ? pr.emboss : pr.AddNewEmboss();
                emboss.val = value;//(value ? ST_OnOff.True : ST_OnOff.False);
            }

        }



        [Obsolete]
        public void SetStrike(bool value)
        {
            CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
            CT_OnOff strike = pr.IsSetStrike() ? pr.strike : pr.AddNewStrike();
            strike.val = value;
        }

        /**
         * Specifies the alignment which shall be applied to the contents of this
         * run.in relation to the default appearance of the run.s text.
         * This allows the text to be repositioned as subscript or superscript without
         * altering the font size of the run.properties.
         *
         * @return VerticalAlign
         * @see VerticalAlign all possible value that could be Applyed to this run
         */
        public VerticalAlign Subscript
        {
            get
            {
                CT_RPr pr = run.rPr;
                return (pr != null && pr.IsSetVertAlign()) ?
                    EnumConverter.ValueOf<VerticalAlign, ST_VerticalAlignRun>(pr.vertAlign.val) : VerticalAlign.BASELINE;
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_VerticalAlignRun ctValign = pr.IsSetVertAlign() ? pr.vertAlign : pr.AddNewVertAlign();
                ctValign.val = EnumConverter.ValueOf<ST_VerticalAlignRun, VerticalAlign>(value);
            }
        }

        public int Kerning
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetKern())
                    return 0;
                return (int)pr.kern.val;
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_HpsMeasure kernmes = pr.IsSetKern() ? pr.kern : pr.AddNewKern();
                kernmes.val = (ulong)value;
            }

        }

        public int CharacterSpacing
        {
            get
            {
                CT_RPr pr = run.rPr;
                if (pr == null || !pr.IsSetSpacing())
                    return 0;
                return int.Parse(pr.spacing.val);
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_SignedTwipsMeasure spc = pr.IsSetSpacing() ? pr.spacing : pr.AddNewSpacing();
                spc.val = value.ToString();
            }
        }

        /**
         * Specifies the fonts which shall be used to display the text contents of
         * this run. Specifies a font which shall be used to format all characters
         * in the ASCII range (0 - 127) within the parent run
         *
         * @return a string representing the font family
         */
        public String FontFamily
        {
            get
            {
                return GetFontFamily(FontCharRange.None);
            }
            set
            {
                SetFontFamily(value, FontCharRange.None);
            }
        }

        public string FontName
        {
            get { return FontFamily; }
        }
        /**
         * Gets the font family for the specified font char range.
         * If fcr is null, the font char range "ascii" is used
         *
         * @param fcr the font char range, defaults to "ansi"
         * @return  a string representing the font famil
         */
        public String GetFontFamily(FontCharRange fcr)
        {
            CT_RPr pr = run.rPr;
            if (pr == null || !pr.IsSetRFonts()) return null;

            CT_Fonts fonts = pr.rFonts;
            switch (fcr == FontCharRange.None ? FontCharRange.Ascii : fcr)
            {
                default:
                case FontCharRange.Ascii:
                    return fonts.ascii;
                case FontCharRange.CS:
                    return fonts.cs;
                case FontCharRange.EastAsia:
                    return fonts.eastAsia;
                case FontCharRange.HAnsi:
                    return fonts.hAnsi;
            }
        }

        /**
         * Specifies the fonts which shall be used to display the text contents of
         * this run. The default handling for fcr == null is to overwrite the
         * ascii font char range with the given font family and also set all not
         * specified font ranges
         *
         * @param fontFamily
         * @param fcr FontCharRange or null for default handling
         */
        public void SetFontFamily(String fontFamily, FontCharRange fcr)
        {
            CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
            CT_Fonts fonts = pr.IsSetRFonts() ? pr.rFonts : pr.AddNewRFonts();

            if (fcr == FontCharRange.None)
            {
                fonts.ascii = (fontFamily);
                if (!fonts.IsSetHAnsi())
                {
                    fonts.hAnsi = (fontFamily);
                }
                if (!fonts.IsSetCs())
                {
                    fonts.cs = (fontFamily);
                }
                if (!fonts.IsSetEastAsia())
                {
                    fonts.eastAsia = (fontFamily);
                }
            }
            else
            {
                switch (fcr)
                {
                    case FontCharRange.Ascii:
                        fonts.ascii = (fontFamily);
                        break;
                    case FontCharRange.CS:
                        fonts.cs = (fontFamily);
                        break;
                    case FontCharRange.EastAsia:
                        fonts.eastAsia = (fontFamily);
                        break;
                    case FontCharRange.HAnsi:
                        fonts.hAnsi = (fontFamily);
                        break;
                }
            }
        }

        /**
         * Specifies the font size which shall be applied to all non complex script
         * characters in the contents of this run.when displayed.
         *
         * @return value representing the font size
         */
        public int FontSize
        {
            get
            {
                CT_RPr pr = run.rPr;
                return (pr != null && pr.IsSetSz()) ? (int)pr.sz.val / 2 : -1;
            }
            set
            {
                CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
                CT_HpsMeasure ctSize = pr.IsSetSz() ? pr.sz : pr.AddNewSz();
                ctSize.val = (ulong)value * 2;
            }
        }

        /**
         * This element specifies the amount by which text shall be raised or
         * lowered for this run.in relation to the default baseline of the
         * surrounding non-positioned text. This allows the text to be repositioned
         * without altering the font size of the contents.
         *
         * @return a big integer representing the amount of text shall be "moved"
         */
        public int GetTextPosition()
        {
            CT_RPr pr = run.rPr;
            return (pr != null && pr.IsSetPosition()) ? int.Parse(pr.position.val)
                    : -1;
        }

        /**
         * This element specifies the amount by which text shall be raised or
         * lowered for this run.in relation to the default baseline of the
         * surrounding non-positioned text. This allows the text to be repositioned
         * without altering the font size of the contents.
         * 
         * If the val attribute is positive, then the parent run.shall be raised
         * above the baseline of the surrounding text by the specified number of
         * half-points. If the val attribute is negative, then the parent run.shall
         * be lowered below the baseline of the surrounding text by the specified
         * number of half-points.
         *         * 
         * If this element is not present, the default value is to leave the
         * formatting applied at previous level in the style hierarchy. If this
         * element is never applied in the style hierarchy, then the text shall not
         * be raised or lowered relative to the default baseline location for the
         * contents of this run.
         */
        public void SetTextPosition(int val)
        {
            CT_RPr pr = run.IsSetRPr() ? run.rPr : run.AddNewRPr();
            CT_SignedHpsMeasure position = pr.IsSetPosition() ? pr.position : pr.AddNewPosition();
            position.val = (val.ToString());
        }

        /**
         * 
         */
        public void RemoveBreak()
        {
            // TODO
        }

        /**
         * Specifies that a break shall be placed at the current location in the run
         * content. 
         * A break is a special character which is used to override the
         * normal line breaking that would be performed based on the normal layout
         * of the document's contents. 
         * @see #AddCarriageReturn() 
         */
        public void AddBreak()
        {
            run.AddNewBr();
        }

        /**
         * Specifies that a break shall be placed at the current location in the run
         * content.
         * A break is a special character which is used to override the
         * normal line breaking that would be performed based on the normal layout
         * of the document's contents.
         * <p>
         * The behavior of this break character (the
         * location where text shall be restarted After this break) shall be
         * determined by its type values.
         * </p>
         * @see BreakType
         */
        public void AddBreak(BreakType type)
        {
            CT_Br br = run.AddNewBr();
            br.type = EnumConverter.ValueOf<ST_BrType, BreakType>(type);
        }

        /**
         * Specifies that a break shall be placed at the current location in the run
         * content. A break is a special character which is used to override the
         * normal line breaking that would be performed based on the normal layout
         * of the document's contents.
         * <p>
         * The behavior of this break character (the
         * location where text shall be restarted After this break) shall be
         * determined by its type (in this case is BreakType.TEXT_WRAPPING as default) and clear attribute values.
         * </p>
         * @see BreakClear
         */
        public void AddBreak(BreakClear Clear)
        {
            CT_Br br = run.AddNewBr();
            br.type = EnumConverter.ValueOf<ST_BrType, BreakType>(BreakType.TEXTWRAPPING);
            br.clear = EnumConverter.ValueOf<ST_BrClear, BreakClear>(Clear);
        }

        /**
         * Specifies that a tab shall be placed at the current location in 
         *  the run content.
         */
        public void AddTab()
        {
            run.AddNewTab();
        }

        public void RemoveTab()
        {
            //TODO
        }

        /**
         * Specifies that a carriage return shall be placed at the
         * current location in the run.content.
         * A carriage return is used to end the current line of text in
         * WordProcess.
         * The behavior of a carriage return in run.content shall be
         * identical to a break character with null type and clear attributes, which
         * shall end the current line and find the next available line on which to
         * continue.
         * The carriage return character forced the following text to be
         * restarted on the next available line in the document.
         */
        public void AddCarriageReturn()
        {
            run.AddNewCr();
        }

        public void RemoveCarriageReturn(int i)
        {
            throw new NotImplementedException();
        }

        /**
         * Adds a picture to the run. This method handles
         *  attaching the picture data to the overall file.
         *  
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_EMF
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_WMF
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_PICT
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_JPEG
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_PNG
         * @see NPOI.XWPF.UserModel.Document#PICTURE_TYPE_DIB
         *  
         * @param pictureData The raw picture data
         * @param pictureType The type of the picture, eg {@link Document#PICTURE_TYPE_JPEG}
         * @param width width in EMUs. To convert to / from points use {@link org.apache.poi.util.Units}
         * @param height height in EMUs. To convert to / from points use {@link org.apache.poi.util.Units}
         * @throws NPOI.Openxml4j.exceptions.InvalidFormatException 
         * @throws IOException 
         */
        public XWPFPicture AddPicture(Stream pictureData, int pictureType, String filename, int width, int height)
        {
            XWPFDocument doc = parent.Document;

            // Add the picture + relationship
            String relationId = doc.AddPictureData(pictureData, pictureType);
            XWPFPictureData picData = (XWPFPictureData)doc.GetRelationById(relationId);

            // Create the Drawing entry for it
            CT_Drawing Drawing = run.AddNewDrawing();
            CT_Inline inline = Drawing.AddNewInline();

            // Do the fiddly namespace bits on the inline
            // (We need full control of what goes where and as what)
            //CT_GraphicalObject tmp = new CT_GraphicalObject();
            //String xml =
            //    "<a:graphic xmlns:a=\"" + "http://schemas.openxmlformats.org/drawingml/2006/main" + "\">" +
            //    "<a:graphicData uri=\"" + "http://schemas.openxmlformats.org/drawingml/2006/picture" + "\">" +
            //    "<pic:pic xmlns:pic=\"" + "http://schemas.openxmlformats.org/drawingml/2006/picture" + "\" />" +
            //    "</a:graphicData>" +
            //    "</a:graphic>";
            //inline.Set((xml));

            XmlDocument xmlDoc = new XmlDocument();
            //XmlElement el = xmlDoc.CreateElement("pic", "pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            inline.graphic = new CT_GraphicalObject();
            inline.graphic.graphicData = new CT_GraphicalObjectData();
            inline.graphic.graphicData.uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";


            // Setup the inline
            inline.distT = (0);
            inline.distR = (0);
            inline.distB = (0);
            inline.distL = (0);

            NPOI.OpenXmlFormats.Dml.WordProcessing.CT_NonVisualDrawingProps docPr = inline.AddNewDocPr();
            long id = parent.Document.DrawingIdManager.ReserveNew();
            docPr.id = (uint)(id);
            /* This name is not visible in Word 2010 anywhere. */
            docPr.name = ("Drawing " + id);
            docPr.descr = (filename);

            NPOI.OpenXmlFormats.Dml.WordProcessing.CT_PositiveSize2D extent = inline.AddNewExtent();
            extent.cx = (width);
            extent.cy = (height);

            // Grab the picture object
            NPOI.OpenXmlFormats.Dml.Picture.CT_Picture pic = new OpenXmlFormats.Dml.Picture.CT_Picture();

            // Set it up
            NPOI.OpenXmlFormats.Dml.Picture.CT_PictureNonVisual nvPicPr = pic.AddNewNvPicPr();

            NPOI.OpenXmlFormats.Dml.CT_NonVisualDrawingProps cNvPr = nvPicPr.AddNewCNvPr();
            /* use "0" for the id. See ECM-576, 20.2.2.3 */
            cNvPr.id = (0);
            /* This name is not visible in Word 2010 anywhere */
            cNvPr.name = ("Picture " + id);
            cNvPr.descr = (filename);

            CT_NonVisualPictureProperties cNvPicPr = nvPicPr.AddNewCNvPicPr();
            cNvPicPr.AddNewPicLocks().noChangeAspect = true;

            CT_BlipFillProperties blipFill = pic.AddNewBlipFill();
            CT_Blip blip = blipFill.AddNewBlip();
            blip.embed = (picData.GetPackageRelationship().Id);
            blipFill.AddNewStretch().AddNewFillRect();

            CT_ShapeProperties spPr = pic.AddNewSpPr();
            CT_Transform2D xfrm = spPr.AddNewXfrm();

            CT_Point2D off = xfrm.AddNewOff();
            off.x = (0);
            off.y = (0);

            NPOI.OpenXmlFormats.Dml.CT_PositiveSize2D ext = xfrm.AddNewExt();
            ext.cx = (width);
            ext.cy = (height);

            CT_PresetGeometry2D prstGeom = spPr.AddNewPrstGeom();
            prstGeom.prst = (ST_ShapeType.rect);
            prstGeom.AddNewAvLst();

            using (var ms = new MemoryStream())
            {
                StreamWriter sw = new StreamWriter(ms);
                pic.Write(sw, "pic:pic");
                sw.Flush();
                ms.Position = 0;
                var sr = new StreamReader(ms);
                var picXml = sr.ReadToEnd();
                inline.graphic.graphicData.AddPicElement(picXml);
            }
            // Finish up
            XWPFPicture xwpfPicture = new XWPFPicture(pic, this);
            pictures.Add(xwpfPicture);
            return xwpfPicture;

        }

        /**
         * Returns the embedded pictures of the run. These
         *  are pictures which reference an external, 
         *  embedded picture image such as a .png or .jpg
         */
        public List<XWPFPicture> GetEmbeddedPictures()
        {
            return pictures;
        }


        /**
         * Add the xml:spaces="preserve" attribute if the string has leading or trailing white spaces
         *
         * @param xs    the string to check
         */
        static void preserveSpaces(CT_Text xs)
        {
            String text = xs.Value;
            if (text != null && (text.StartsWith(" ") || text.EndsWith(" ")))
            {
                //    XmlCursor c = xs.NewCursor();
                //    c.ToNextToken();
                //    c.InsertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace", "space"), "preserve");
                //    c.Dispose();
                xs.space = "preserve";
            }
        }

        /**
         * Returns the string version of the text, with tabs and
         *  carriage returns in place of their xml equivalents.
         */
        public override String ToString()
        {
            return Text;
        }
    }

}