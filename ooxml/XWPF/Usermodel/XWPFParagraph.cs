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
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Text;
    using NPOI.XWPF.Util;
    using NPOI.Util;
    using System.Collections;
    /**
     * Sketch of XWPF paragraph class
     */
    public class XWPFParagraph : IBodyElement
    {
        private CT_P paragraph;
        protected IBody part;
        /** For access to the document's hyperlink, comments, tables etc */
        protected XWPFDocument document;
        protected List<XWPFRun> runs;

        private StringBuilder footnoteText = new StringBuilder();

        public XWPFParagraph(CT_P prgrph, IBody part)
        {
            this.paragraph = prgrph;
            this.part = part;

            this.document = part.GetXWPFDocument();

            if (document == null)
            {
                throw new NullReferenceException();
            }
            // Build up the character runs
            runs = new List<XWPFRun>();

            BuildRunsInOrderFromXml(paragraph.Items);
            // Look for bits associated with the runs
            foreach (XWPFRun run in runs)
            {
                CT_R r = run.GetCTR();
                if (document != null)
                {
                    foreach (object o in r.Items)
                    {
                        if (o is CT_FtnEdnRef)
                        {
                            CT_FtnEdnRef ftn = (CT_FtnEdnRef)o;
                            footnoteText.Append("[").Append(ftn.id).Append(": ");
                            
                            XWPFFootnote footnote = document.GetFootnoteByID(int.Parse(ftn.id));
                            if (footnote == null)
                                footnote = document.GetEndnoteByID(int.Parse(ftn.id));
                               //ftn.DomNode.LocalName.Equals("footnoteReference") ?
                               //      document.GetFootnoteByID(ftn.Id.IntValue()) :
                               //      document.GetEndnoteByID(ftn.Id.IntValue());

                            bool first = true;
                            foreach (XWPFParagraph p in footnote.Paragraphs)
                            {
                                if (!first)
                                {
                                    footnoteText.Append("\n");
                                    first = false;
                                }
                                footnoteText.Append(p.Text);
                            }

                            footnoteText.Append("]");
                        }
                    }
                }
            }

            // Get all our child nodes in order, and process them
            //  into XWPFRuns where we can
            /*XmlCursor c = paragraph.NewCursor();
            c.SelectPath("child::*");
            while (c.ToNextSelection()) {
               XmlObject o = c.Object;
               if(o is CT_R) {
                  Runs.Add(new XWPFRun((CT_R)o, this));
               }
               if(o is CT_Hyperlink) {
                  CT_Hyperlink link = (CT_Hyperlink)o;
                  foreach(CTR r in link.RList) {
                     Runs.Add(new XWPFHyperlinkRun(link, r, this));
                  }
               }
               if(o is CT_SdtRun) {
                  CT_SdtContentRun run = ((CT_SdtRun)o).SdtContent;
                  foreach(CTR r in Run.RList) {
                     Runs.Add(new XWPFRun(r, this));
                  }
               }
               if(o is CT_RunTrackChange) {
                  foreach(CTR r in ((CT_RunTrackChange)o).RList) {
                     Runs.Add(new XWPFRun(r, this));
                  }
               }
               if(o is CT_SimpleField) {
                  foreach(CTR r in ((CT_SimpleField)o).RList) {
                     Runs.Add(new XWPFRun(r, this));
                  }
               }
            }*/
        }
        /**
         * Identifies (in order) the parts of the paragraph /
         *  sub-paragraph that correspond to character text
         *  runs, and builds the appropriate runs for these.
         */
        private void BuildRunsInOrderFromXml(ArrayList items)
        {
            foreach (object o in items)
            {
                if (o is CT_R)
                {
                    runs.Add(new XWPFRun((CT_R)o, this));
                }
                if (o is CT_Hyperlink1)
                {
                    CT_Hyperlink1 link = (CT_Hyperlink1)o;
                    foreach (CT_R r in link.GetRList())
                    {
                        runs.Add(new XWPFHyperlinkRun(link, r, this));
                    }
                }
                if (o is CT_SdtRun)
                {
                    CT_SdtContentRun run = ((CT_SdtRun)o).sdtContent;
                    foreach (CT_R r in run.GetRList())
                    {
                        runs.Add(new XWPFRun(r, this));
                    }
                }
                if (o is CT_RunTrackChange)
                {
                    foreach (CT_R r in ((CT_RunTrackChange)o).GetRList())
                    {
                        runs.Add(new XWPFRun(r, this));
                    }
                }
                if (o is CT_SimpleField)
                {
                    foreach (CT_R r in ((CT_SimpleField)o).GetRList())
                    {
                        runs.Add(new XWPFRun(r, this));
                    }
                }
                if (o is CT_SmartTagRun)
                {
                    // Smart Tags can be nested many times. 
                    // This implementation does not preserve the tagging information
                    BuildRunsInOrderFromXml((o as CT_SmartTagRun).Items);
                }
            }
        }


        internal CT_P GetCTP()
        {
            return paragraph;
        }

        public IList<XWPFRun> Runs
        {
			get
			{
				return runs.AsReadOnly();
			}
        }

        public bool IsEmpty
        {
			get
			{
				return paragraph.Items.Count == 0;
			}
        }

        public XWPFDocument Document
        {
			get
			{
				return document;
			}
        }

        /**
         * Return the textual content of the paragraph, including text from pictures
         * in it.
         */
        public String Text
        {
            get
            {
                StringBuilder out1 = new StringBuilder();
                foreach (XWPFRun run in runs)
                {
                    out1.Append(run.ToString());
                }
                out1.Append(footnoteText);
                return out1.ToString();
            }
        }

        /**
         * Return styleID of the paragraph if style exist for this paragraph
         * if not, null will be returned     
         * @return		styleID as String
         */
        public String StyleID
        {
            get
            {
                if (paragraph.pPr != null)
                {
                    if (paragraph.pPr.pStyle != null)
                    {
                        if (paragraph.pPr.pStyle.val != null)
                            return paragraph.pPr.pStyle.val;
                    }
                }
                return null;
            }
        }
        /**
         * If style exist for this paragraph
         * NumId of the paragraph will be returned.
         * If style not exist null will be returned     
         * @return	NumID as Bigint
         */
        public string GetNumID()
        {
            if (paragraph.pPr != null)
            {
                if (paragraph.pPr.numPr != null)
                {
                    if (paragraph.pPr.numPr.numId != null)
                        return paragraph.pPr.numPr.numId.val;
                }
            }
            return null;
        }

        /**
         * SetNumID of Paragraph
         * @param numPos
         */
        public void SetNumID(string numId)
        {
            if (paragraph.pPr == null)
                paragraph.AddNewPPr();
            if (paragraph.pPr.numPr == null)
                paragraph.pPr.AddNewNumPr();
            if (paragraph.pPr.numPr.numId == null)
            {
                paragraph.pPr.numPr.AddNewNumId();
            }
            paragraph.pPr.numPr.ilvl = new CT_DecimalNumber();
            paragraph.pPr.numPr.ilvl.val = "0";
            paragraph.pPr.numPr.numId.val = numId;
        }
        /// <summary>
        /// Set NumID and level of Paragraph
        /// </summary>
        /// <param name="numId"></param>
        /// <param name="ilvl"></param>
        public void SetNumID(string numId, string ilvl)
        {
            if (paragraph.pPr == null)
                paragraph.AddNewPPr();
            if (paragraph.pPr.numPr == null)
                paragraph.pPr.AddNewNumPr();
            if (paragraph.pPr.numPr.numId == null)
            {
                paragraph.pPr.numPr.AddNewNumId();
            }
            paragraph.pPr.numPr.ilvl = new CT_DecimalNumber();
            paragraph.pPr.numPr.ilvl.val = ilvl;
            paragraph.pPr.numPr.numId.val = (numId);
        }
        /**
         * Returns the text of the paragraph, but not of any objects in the
         * paragraph
         */
        public String ParagraphText
        {
            get
            {
                StringBuilder out1 = new StringBuilder();
                foreach (XWPFRun run in runs)
                {
                    out1.Append(run.ToString());
                }
                return out1.ToString();
            }
        }

        /**
         * Returns any text from any suitable pictures in the paragraph
         */
        public String GetPictureText()
        {
            StringBuilder out1 = new StringBuilder();
            foreach(XWPFRun run in runs) {
                out1.Append(run.GetPictureText());
            }
            return out1.ToString();
        }

        /**
         * Returns the footnote text of the paragraph
         *
         * @return  the footnote text or empty string if the paragraph does not have footnotes
         */
        public String FootnoteText
        {
			get
			{
				return footnoteText.ToString();
			}
        }

        /// <summary>
        /// Appends a new run to this paragraph
        /// </summary>
        /// <returns>a new text run</returns>
        public XWPFRun CreateRun()
        {
            XWPFRun xwpfRun = new XWPFRun(paragraph.AddNewR(), this);
            runs.Add(xwpfRun);
            return xwpfRun;
        }

        /**
         * Returns the paragraph alignment which shall be applied to text in this
         * paragraph.
         * <p>
         * If this element is not Set on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then no alignment is applied to the
         * paragraph.
         * </p>
         *
         * @return the paragraph alignment of this paragraph.
         */
        public ParagraphAlignment GetAlignment()
        {
            CT_PPr pr = GetCTPPr();
            return pr == null || !pr.IsSetJc() ? ParagraphAlignment.LEFT : EnumConverter.ValueOf<ParagraphAlignment, ST_Jc>(pr.jc.val);
        }

        /**
         * Specifies the paragraph alignment which shall be applied to text in this
         * paragraph.
         * <p>
         * <p>
         * If this element is not Set on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then no alignment is applied to the
         * paragraph.
         * </p>
         *
         * @param align the paragraph alignment to apply to this paragraph.
         */
        public void SetAlignment(ParagraphAlignment align)
        {
            CT_PPr pr = GetCTPPr();
            CT_Jc jc = pr.IsSetJc() ? pr.jc : pr.AddNewJc();
            jc.val = EnumConverter.ValueOf<ST_Jc, ParagraphAlignment>(align);
        }

        /**
         * Returns the text vertical alignment which shall be applied to text in
         * this paragraph.
         * <p>
         * If the line height (before any Added spacing) is larger than one or more
         * characters on the line, all characters will be aligned to each other as
         * specified by this element.
         * </p>
         * <p>
         * If this element is omitted on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then the vertical alignment of all
         * characters on the line shall be automatically determined by the consumer.
         * </p>
         *
         * @return the vertical alignment of this paragraph.
         */
        public TextAlignment GetVerticalAlignment()
        {
            CT_PPr pr = GetCTPPr();
            return (pr == null || !pr.IsSetTextAlignment()) ? TextAlignment.AUTO
                    : EnumConverter.ValueOf<TextAlignment, ST_TextAlignment>(pr.textAlignment.val);
        }

        /**
         * Specifies the text vertical alignment which shall be applied to text in
         * this paragraph.
         * <p>
         * If the line height (before any Added spacing) is larger than one or more
         * characters on the line, all characters will be aligned to each other as
         * specified by this element.
         * </p>
         * <p>
         * If this element is omitted on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then the vertical alignment of all
         * characters on the line shall be automatically determined by the consumer.
         * </p>
         *
         * @param valign the paragraph vertical alignment to apply to this
         *               paragraph.
         */
        public void SetVerticalAlignment(TextAlignment valign)
        {
            CT_PPr pr = GetCTPPr();
            CT_TextAlignment textAlignment = pr.IsSetTextAlignment() ? pr
                    .textAlignment : pr.AddNewTextAlignment();
            //STTextAlignment.Enum en = STTextAlignment.Enum
            //        .forInt(valign.Value);
            textAlignment.val = EnumConverter.ValueOf<ST_TextAlignment, TextAlignment>(valign);
        }

        /**
         * Specifies the border which shall be displayed above a Set of paragraphs
         * which have the same Set of paragraph border Settings.
         * <p>
         * <p>
         * To determine if any two adjoining paragraphs shall have an individual top
         * and bottom border or a between border, the Set of borders on the two
         * adjoining paragraphs are Compared. If the border information on those two
         * paragraphs is identical for all possible paragraphs borders, then the
         * between border is displayed. Otherwise, the paragraph shall use its
         * bottom border and the following paragraph shall use its top border,
         * respectively. If this border specifies a space attribute, that value
         * determines the space above the text (ignoring any spacing above) which
         * should be left before this border is Drawn, specified in points.
         * </p>
         * <p>
         * If this element is omitted on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then no between border shall be applied
         * above identical paragraphs.
         * </p>
         * <b>This border can only be a line border.</b>
         *
         * @param border
         * @see Borders for a list of all types of borders
         */
        public void SetBorderTop(Borders border)
        {
            CT_PBdr ct = GetCTPBrd(true);

            CT_Border pr = (ct != null && ct.IsSetTop()) ? ct.top : ct.AddNewTop();
            if (border == Borders.NONE)
                ct.UnsetTop();
            else
                pr.val = EnumConverter.ValueOf<ST_Border, Borders>(border);
        }

        /**
         * Specifies the border which shall be displayed above a Set of paragraphs
         * which have the same Set of paragraph border Settings.
         *
         * @return paragraphBorder - the top border for the paragraph
         * @see #setBorderTop(Borders)
         * @see Borders a list of all types of borders
         */
        public Borders GetBorderTop()
        {
            CT_PBdr border = GetCTPBrd(false);
            CT_Border ct = null;
            if (border != null)
            {
                ct = border.top;
            }
            ST_Border ptrn = (ct != null) ? ct.val : ST_Border.none;
            return EnumConverter.ValueOf<Borders, ST_Border>(ptrn);
        }

        /**
 * 
 * <p>

 * </p>
 * <p>
 * If this element is omitted on a given paragraph, its value is determined
 * by the Setting previously Set at any level of the style hierarchy (i.e.
 * that previous Setting remains unChanged). If this Setting is never
 * specified in the style hierarchy, then no between border shall be applied
 * below identical paragraphs.
 * </p>
 * <b>This border can only be a line border.</b>
 */
        /// <summary>
        /// Specifies the border which shall be displayed below a Set of paragraphs
        /// which have the same Set of paragraph border Settings.
        /// </summary>
        /// <param name="border">a list of all types of borders</param>
        /// <remarks>
        ///  To determine if any two adjoining paragraphs shall have an individual top
        /// and bottom border or a between border, the Set of borders on the two
        /// adjoining paragraphs are Compared. If the border information on those two
        /// paragraphs is identical for all possible paragraphs borders, then the
        /// between border is displayed. Otherwise, the paragraph shall use its
        /// bottom border and the following paragraph shall use its top border,
        /// respectively. If this border specifies a space attribute, that value
        /// determines the space After the bottom of the text (ignoring any space
        /// below) which should be left before this border is Drawn, specified in
        /// points.
        /// </remarks>
        public void SetBorderBottom(Borders border)
        {
            CT_PBdr ct = GetCTPBrd(true);
            CT_Border pr = ct.IsSetBottom() ? ct.bottom : ct.AddNewBottom();
            if (border == Borders.NONE)
                ct.UnsetBottom();
            else
                pr.val = EnumConverter.ValueOf<ST_Border, Borders>(border);
        }

        /// <summary>
        ///Specifies the border which shall be displayed below a Set of
        /// paragraphs which have the same Set of paragraph border Settings.
        /// </summary>
        /// <returns>the bottom border for the paragraph</returns>
        public Borders GetBorderBottom()
        {
            CT_PBdr border = GetCTPBrd(false);
            CT_Border ct = null;
            if (border != null)
            {
                ct = border.bottom;
            }
            ST_Border ptrn = ct != null ? ct.val : ST_Border.none;
            return EnumConverter.ValueOf<Borders, ST_Border>(ptrn);

        }

        /**
         * Specifies the border which shall be displayed on the left side of the
         * page around the specified paragraph.
         * <p>
         * To determine if any two adjoining paragraphs should have a left border
         * which spans the full line height or not, the left border shall be Drawn
         * between the top border or between border at the top (whichever would be
         * rendered for the current paragraph), and the bottom border or between
         * border at the bottom (whichever would be rendered for the current
         * paragraph).
         * </p>
         * <p>
         * If this element is omitted on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then no left border shall be applied.
         * </p>
         * <b>This border can only be a line border.</b>
         *
         * @param border
         * @see Borders for a list of all possible borders
         */
        public void SetBorderLeft(Borders border)
        {
            CT_PBdr ct = GetCTPBrd(true);
            CT_Border pr = ct.IsSetLeft() ? ct.left : ct.AddNewLeft();
            if (border == Borders.NONE)
                ct.UnsetLeft();
            else
                pr.val = EnumConverter.ValueOf<ST_Border, Borders>(border);

        }

        /// <summary>
        /// Specifies the border which shall be displayed on the left side of the
        /// page around the specified paragraph.
        /// </summary>
        /// <returns>the left border for the paragraph</returns>
        public Borders GetBorderLeft()
        {
            CT_PBdr border = GetCTPBrd(false);
            CT_Border ct = null;
            if (border != null)
            {
                ct = border.left;
            }
            ST_Border ptrn = ct != null ? ct.val : ST_Border.none;
            return EnumConverter.ValueOf<Borders, ST_Border>(ptrn);
        }

        /**
         * Specifies the border which shall be displayed on the right side of the
         * page around the specified paragraph.
         * <p>
         * To determine if any two adjoining paragraphs should have a right border
         * which spans the full line height or not, the right border shall be Drawn
         * between the top border or between border at the top (whichever would be
         * rendered for the current paragraph), and the bottom border or between
         * border at the bottom (whichever would be rendered for the current
         * paragraph).
         * </p>
         * <p>
         * If this element is omitted on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then no right border shall be applied.
         * </p>
         * <b>This border can only be a line border.</b>
         *
         * @param border
         * @see Borders for a list of all possible borders
         */
        public void SetBorderRight(Borders border)
        {
            CT_PBdr ct = GetCTPBrd(true);
            CT_Border pr = ct.IsSetRight() ? ct.right : ct.AddNewRight();
            if (border == Borders.NONE)
                ct.UnsetRight();
            else
                pr.val = EnumConverter.ValueOf<ST_Border, Borders>(border);
        }

        /**
         * Specifies the border which shall be displayed on the right side of the
         * page around the specified paragraph.
         *
         * @return ParagraphBorder - the right border for the paragraph
         * @see #setBorderRight(Borders)
         * @see Borders for a list of all possible borders
         */
        public Borders GetBorderRight()
        {
            CT_PBdr border = GetCTPBrd(false);
            CT_Border ct = null;
            if (border != null)
            {
                ct = border.right;
            }
            ST_Border ptrn = ct != null ? ct.val : ST_Border.none;
            return EnumConverter.ValueOf<Borders, ST_Border>(ptrn);
        }

        /**
         * Specifies the border which shall be displayed between each paragraph in a
         * Set of paragraphs which have the same Set of paragraph border Settings.
         * <p>
         * To determine if any two adjoining paragraphs should have a between border
         * or an individual top and bottom border, the Set of borders on the two
         * adjoining paragraphs are Compared. If the border information on those two
         * paragraphs is identical for all possible paragraphs borders, then the
         * between border is displayed. Otherwise, each paragraph shall use its
         * bottom and top border, respectively. If this border specifies a space
         * attribute, that value is ignored - this border is always located at the
         * bottom of each paragraph with an identical following paragraph, taking
         * into account any space After the line pitch.
         * </p>
         * <p>
         * If this element is omitted on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then no between border shall be applied
         * between identical paragraphs.
         * </p>
         * <b>This border can only be a line border.</b>
         *
         * @param border
         * @see Borders for a list of all possible borders
         */
        public void SetBorderBetween(Borders border)
        {
            CT_PBdr ct = GetCTPBrd(true);
            CT_Border pr = ct.IsSetBetween() ? ct.between : ct.AddNewBetween();
            if (border == Borders.NONE)
                ct.UnsetBetween();
            else
                pr.val = EnumConverter.ValueOf<ST_Border, Borders>(border);
        }

        /**
         * Specifies the border which shall be displayed between each paragraph in a
         * Set of paragraphs which have the same Set of paragraph border Settings.
         *
         * @return ParagraphBorder - the between border for the paragraph
         * @see #setBorderBetween(Borders)
         * @see Borders for a list of all possible borders
         */
        public Borders GetBorderBetween()
        {
            CT_PBdr border = GetCTPBrd(false);
            CT_Border ct = null;
            if (border != null)
            {
                ct = border.between;
            }
            ST_Border ptrn = ct != null ? ct.val : ST_Border.none;
            return EnumConverter.ValueOf<Borders, ST_Border>(ptrn);
        }

        /**
         * Specifies that when rendering this document in a paginated
         * view, the contents of this paragraph are rendered on the start of a new
         * page in the document.
         * <p>
         * If this element is omitted on a given paragraph,
         * its value is determined by the Setting previously Set at any level of the
         * style hierarchy (i.e. that previous Setting remains unChanged). If this
         * Setting is never specified in the style hierarchy, then this property
         * shall not be applied. Since the paragraph is specified to start on a new
         * page, it begins page two even though it could have fit on page one.
         * </p>
         *
         * @param pageBreak -
         *                  bool value
         */
        public void SetPageBreak(bool pageBreak)
        {
            CT_PPr ppr = GetCTPPr();
            CT_OnOff ct_pageBreak = ppr.IsSetPageBreakBefore() ? ppr
                    .pageBreakBefore : ppr.AddNewPageBreakBefore();
                ct_pageBreak.val =pageBreak;
        }

        /**
         * Specifies that when rendering this document in a paginated
         * view, the contents of this paragraph are rendered on the start of a new
         * page in the document.
         * <p>
         * If this element is omitted on a given paragraph,
         * its value is determined by the Setting previously Set at any level of the
         * style hierarchy (i.e. that previous Setting remains unChanged). If this
         * Setting is never specified in the style hierarchy, then this property
         * shall not be applied. Since the paragraph is specified to start on a new
         * page, it begins page two even though it could have fit on page one.
         * </p>
         *
         * @return bool - if page break is Set
         */
        public bool IsPageBreak
        {
            get
            {
                CT_PPr ppr = GetCTPPr();
                CT_OnOff ct_pageBreak = ppr.IsSetPageBreakBefore() ? ppr
                        .pageBreakBefore : null;
                if (ct_pageBreak != null
                        && ct_pageBreak.val)
                {
                    return true;
                }
                return false;
            }
        }
        
        /**
         * Specifies the spacing that should be Added After the last line in this
         * paragraph in the document in absolute units.
         *
         * @return int - value representing the spacing After the paragraph
         */
        public int SpacingAfter
        {
            get
            {
                CT_Spacing spacing = GetCTSpacing(false);
                return (spacing != null && spacing.IsSetAfter()) ? (int)spacing.after : -1;
            }
            set 
            {
                CT_Spacing spacing = GetCTSpacing(true);
                if (spacing != null)
                {
                    //BigInteger bi = new BigInteger(spaces);
                    spacing.after = (ulong)value;
                }
            }
        }

        /**
         * Specifies the spacing that should be Added After the last line in this
         * paragraph in the document in line units.
         * <b>The value of this attribute is
         * specified in one hundredths of a line.
         * </b>
         * <p>
         * If the AfterAutoSpacing attribute
         * is also specified, then this attribute value is ignored. If this Setting
         * is never specified in the style hierarchy, then its value shall be zero
         * (if needed)
         * </p>
         *
         * @param spaces -
         *               a positive whole number, whose contents consist of a
         *               measurement in twentieths of a
         */
        public void SetSpacingAfterLines(int spaces)
        {
            CT_Spacing spacing = GetCTSpacing(true);
            //BigInteger bi = new BigInteger("" + spaces);
            spacing.afterLines = (spaces.ToString());
        }


        /**
         * Specifies the spacing that should be Added After the last line in this
         * paragraph in the document in absolute units.
         *
         * @return bigint - value representing the spacing After the paragraph
         * @see #setSpacingAfterLines(int)
         */
        public int GetSpacingAfterLines()
        {
            CT_Spacing spacing = GetCTSpacing(false);
            return (spacing != null && spacing.IsSetAfterLines()) ? int.Parse(spacing.afterLines) : -1;
        }


        /**
         * Specifies the spacing that should be Added above the first line in this
         * paragraph in the document in absolute units.
         * <p>
         * If the beforeLines attribute or the beforeAutoSpacing attribute is also
         * specified, then this attribute value is ignored.
         * </p>
         *
         * @param spaces
         */
        public void SetSpacingBefore(int spaces)
        {
            CT_Spacing spacing = GetCTSpacing(true);
            //BigInteger bi = new BigInteger("" + spaces);
            spacing.before = (ulong)spaces;
        }

        /**
         * Specifies the spacing that should be Added above the first line in this
         * paragraph in the document in absolute units.
         *
         * @return the spacing that should be Added above the first line
         * @see #setSpacingBefore(int)
         */
        public int GetSpacingBefore()
        {
            CT_Spacing spacing = GetCTSpacing(false);
            return (spacing != null && spacing.IsSetBefore()) ? (int)spacing.before : -1;
        }

        /**
         * Specifies the spacing that should be Added before the first line in this
         * paragraph in the document in line units. <b> The value of this attribute
         * is specified in one hundredths of a line. </b>
         * <p>
         * If the beforeAutoSpacing attribute is also specified, then this attribute
         * value is ignored. If this Setting is never specified in the style
         * hierarchy, then its value shall be zero.
         * </p>
         *
         * @param spaces
         */
        public void SetSpacingBeforeLines(int spaces)
        {
            CT_Spacing spacing = GetCTSpacing(true);
            //BigInteger bi = new BigInteger("" + spaces);
            spacing.beforeLines = spaces.ToString();
        }

        /**
         * Specifies the spacing that should be Added before the first line in this paragraph in the
         * document in line units.
         * The value of this attribute is specified in one hundredths of a line.
         *
         * @return the spacing that should be Added before the first line in this paragraph
         * @see #setSpacingBeforeLines(int)
         */
        public int GetSpacingBeforeLines()
        {
            CT_Spacing spacing = GetCTSpacing(false);
            return (spacing != null && spacing.IsSetBeforeLines()) ? int.Parse(spacing.beforeLines) : -1;
        }


        /**
         * Specifies how the spacing between lines is calculated as stored in the
         * line attribute. If this attribute is omitted, then it shall be assumed to
         * be of a value auto if a line attribute value is present.
         *
         * @param rule
         * @see LineSpacingRule
         */
        public void SetSpacingLineRule(LineSpacingRule rule)
        {
            CT_Spacing spacing = GetCTSpacing(true);
            spacing.lineRule = EnumConverter.ValueOf<ST_LineSpacingRule, LineSpacingRule>(rule);
        }

        /**
         * Specifies how the spacing between lines is calculated as stored in the
         * line attribute. If this attribute is omitted, then it shall be assumed to
         * be of a value auto if a line attribute value is present.
         *
         * @return rule
         * @see LineSpacingRule
         * @see #setSpacingLineRule(LineSpacingRule)
         */
        public LineSpacingRule GetSpacingLineRule()
        {
            CT_Spacing spacing = GetCTSpacing(false);
            return (spacing != null && spacing.IsSetLineRule()) ? 
                EnumConverter.ValueOf<LineSpacingRule,ST_LineSpacingRule>(spacing.lineRule) : LineSpacingRule.AUTO;
        }


        /**
         * Specifies the indentation which shall be placed between the left text
         * margin for this paragraph and the left edge of that paragraph's content
         * in a left to right paragraph, and the right text margin and the right
         * edge of that paragraph's text in a right to left paragraph
         * <p>
         * If this attribute is omitted, its value shall be assumed to be zero.
         * Negative values are defined such that the text is Moved past the text margin,
         * positive values Move the text inside the text margin.
         * </p>
         *
         * @return indentation or null if indentation is not Set
         */
        public int IndentationLeft
        {
            get
            {
                CT_Ind indentation = GetCTInd(false);
                return (indentation != null && indentation.IsSetLeft()) ? int.Parse(indentation.left)
                        : -1;
            }
            set 
            {
                CT_Ind indent = GetCTInd(true);
                //BigInteger bi = new BigInteger("" + indentation);
                indent.left = value.ToString();            
            }
        }

        /**
         * Specifies the indentation which shall be placed between the right text
         * margin for this paragraph and the right edge of that paragraph's content
         * in a left to right paragraph, and the right text margin and the right
         * edge of that paragraph's text in a right to left paragraph
         * <p>
         * If this attribute is omitted, its value shall be assumed to be zero.
         * Negative values are defined such that the text is Moved past the text margin,
         * positive values Move the text inside the text margin.
         * </p>
         *
         * @return indentation or null if indentation is not Set
         */

        public int IndentationRight
        {
            get
            {
                CT_Ind indentation = GetCTInd(false);
                return (indentation != null && indentation.IsSetRight()) ? int.Parse(indentation.right)
                        : -1;
            }
            set 
            {
                CT_Ind indent = GetCTInd(true);
                //BigInteger bi = new BigInteger("" + indentation);
                indent.right = value.ToString();
            }
        }

        /**
         * Specifies the indentation which shall be Removed from the first line of
         * the parent paragraph, by moving the indentation on the first line back
         * towards the beginning of the direction of text flow.
         * This indentation is
         * specified relative to the paragraph indentation which is specified for
         * all other lines in the parent paragraph.
         * The firstLine and hanging
         * attributes are mutually exclusive, if both are specified, then the
         * firstLine value is ignored.
         *
         * @return indentation or null if indentation is not Set
         */
        public int IndentationHanging
        {
            get
            {
                CT_Ind indentation = GetCTInd(false);
                return (indentation != null && indentation.IsSetHanging()) ? (int)indentation.hanging : -1;
            }
            set 
            {
                CT_Ind indent = GetCTInd(true);
                //BigInteger bi = new BigInteger("" + indentation);
                indent.hanging = (ulong)value;            
            }
        }


        /**
         * Specifies the Additional indentation which shall be applied to the first
         * line of the parent paragraph. This Additional indentation is specified
         * relative to the paragraph indentation which is specified for all other
         * lines in the parent paragraph.
         * The firstLine and hanging attributes are
         * mutually exclusive, if both are specified, then the firstLine value is
         * ignored.
         * If the firstLineChars attribute is also specified, then this
         * value is ignored.
         * If this attribute is omitted, then its value shall be
         * assumed to be zero (if needed).
         *
         * @return indentation or null if indentation is not Set
         */
        public int IndentationFirstLine
        {
            get
            {
                CT_Ind indentation = GetCTInd(false);
                return (indentation != null && indentation.IsSetFirstLine()) ? (int)indentation.firstLine
                        : -1;
            }
            set 
            {
                CT_Ind indent = GetCTInd(true);
                //BigInteger bi = new BigInteger("" + indentation);
                indent.firstLine = (long)value;
            }
        }


        /**
         * This element specifies whether a consumer shall break Latin text which
         * exceeds the text extents of a line by breaking the word across two lines
         * (breaking on the character level) or by moving the word to the following
         * line (breaking on the word level).
         *
         * @return bool
         */
        public bool IsWordWrap
        {
            get
            {
                CT_OnOff wordWrap = GetCTPPr().IsSetWordWrap() ? GetCTPPr()
                        .wordWrap : null;
                if (wordWrap != null)
                {
                    return wordWrap.val;
                }
                return false;
            }
            set
            {
                CT_OnOff wordWrap = GetCTPPr().IsSetWordWrap() ? GetCTPPr()
            .wordWrap : GetCTPPr().AddNewWordWrap();
                if (value)
                    wordWrap.val = true;
                else
                    wordWrap.UnSetVal();
            }
        }

        /**
         * @return  the style of the paragraph
         */
        public String Style
        {
            get
            {
                CT_PPr pr = GetCTPPr();
                CT_String style = pr.IsSetPStyle() ? pr.pStyle : null;
                return style != null ? style.val : null;
            }
            set 
            {
                CT_PPr pr = GetCTPPr();
                CT_String style = pr.pStyle != null ? pr.pStyle : pr.AddNewPStyle();
                style.val = value;
            }
        }

        /**
         * Get a <b>copy</b> of the currently used CTPBrd, if none is used, return
         * a new instance.
         */
        private CT_PBdr GetCTPBrd(bool create)
        {
            CT_PPr pr = GetCTPPr();
            CT_PBdr ct = pr.IsSetPBdr() ? pr.pBdr : null;
            if (create && ct == null)
                ct = pr.AddNewPBdr();
            return ct;

        }

        /**
         * Get a <b>copy</b> of the currently used CTSpacing, if none is used,
         * return a new instance.
         */
        private CT_Spacing GetCTSpacing(bool create)
        {
            CT_PPr pr = GetCTPPr();
            CT_Spacing ct = pr.spacing == null ? null : pr.spacing;
            if (create && ct == null)
                ct = pr.AddNewSpacing();
            return ct;
        }

        /**
         * Get a <b>copy</b> of the currently used CTPInd, if none is used, return
         * a new instance.
         */
        private CT_Ind GetCTInd(bool create)
        {
            CT_PPr pr = GetCTPPr();
            CT_Ind ct = pr.ind == null ? null : pr.ind;
            if (create && ct == null)
                ct = pr.AddNewInd();
            return ct;
        }

        /**
         * Get a <b>copy</b> of the currently used CTPPr, if none is used, return
         * a new instance.
         */
        public CT_PPr GetCTPPr()
        {
            CT_PPr pr = paragraph.pPr == null ? paragraph.AddNewPPr()
                    : paragraph.pPr;
            return pr;
        }


        /**
         * add a new run at the end of the position of 
         * the content of parameter run
         * @param run
         */
        protected void AddRun(CT_R Run)
        {
            int pos;
            pos = paragraph.GetRList().Count;
            paragraph.AddNewR();
            paragraph.SetRArray(pos, Run);
        }

        /**
         * this methods parse the paragraph and search for the string searched.
         * If it Finds the string, it will return true and the position of the String
         * will be saved in the parameter startPos.
         * @param searched
         * @param startPos
         */
        public TextSegement SearchText(String searched, PositionInParagraph startPos)
        {

            int startRun = startPos.Run,
                startText = startPos.Text,
                startChar = startPos.Char;
            int beginRunPos = 0, candCharPos = 0;
            bool newList = false;
            for (int RunPos = startRun; RunPos < paragraph.GetRList().Count; RunPos++)
            {
                int beginTextPos = 0, beginCharPos = 0, textPos = 0, charPos = 0;
                CT_R ctRun = paragraph.GetRList()[RunPos];
                foreach(object o in ctRun.Items)
                {
                    if (o is CT_Text)
                    {
                        if (textPos >= startText)
                        {
                            String candidate = ((CT_Text)o).Value;
                            if (RunPos == startRun)
                                charPos = startChar;
                            else
                                charPos = 0;
                            for (; charPos < candidate.Length; charPos++)
                            {
                                if ((candidate[charPos] == searched[0]) && (candCharPos == 0))
                                {
                                    beginTextPos = textPos;
                                    beginCharPos = charPos;
                                    beginRunPos = RunPos;
                                    newList = true;
                                }
                                if (candidate[charPos] == searched[candCharPos])
                                {
                                    if (candCharPos + 1 < searched.Length)
                                        candCharPos++;
                                    else if (newList)
                                    {
                                        TextSegement segement = new TextSegement();
                                        segement.BeginRun = (beginRunPos);
                                        segement.BeginText = (beginTextPos);
                                        segement.BeginChar = (beginCharPos);
                                        segement.EndRun = (RunPos);
                                        segement.EndText = (textPos);
                                        segement.EndChar = (charPos);
                                        return segement;
                                    }
                                }
                                else
                                    candCharPos = 0;
                            }
                        }
                        textPos++;
                    }
                    else if (o is CT_ProofErr)
                    {
                        //c.RemoveXml();
                    }
                    else if (o is CT_RPr)
                    {
                        //do nothing
                    }
                    else
                        candCharPos = 0;
                }
            }
            return null;
        }

        /**
         * insert a new Run in RunArray
         * @param pos
         * @return  the inserted run
         */
        public XWPFRun InsertNewRun(int pos)
        {
            if (pos >= 0 && pos <= paragraph.SizeOfRArray())
            {
                CT_R ctRun = paragraph.InsertNewR(pos);
                XWPFRun newRun = new XWPFRun(ctRun, this);
                runs.Insert(pos, newRun);
                return newRun;
            }
            return null;
        }



        /**
         * Get a Text
         * @param segment
         */
        public String GetText(TextSegement segment)
        {
            //int RunBegin = segment.BeginRun;
            //int textBegin = segment.BeginText;
            //int charBegin = segment.BeginChar;
            //int RunEnd = segment.EndRun;
            //int textEnd = segment.EndText;
            //int charEnd = segment.EndChar;
            //StringBuilder out1 = new StringBuilder();
            //for (int i = RunBegin; i <= RunEnd; i++)
            //{
            //    int startText = 0, endText = paragraph.GetRArray(i).TList.Size() - 1;
            //    if (i == RunBegin)
            //        startText = textBegin;
            //    if (i == RunEnd)
            //        endText = textEnd;
            //    for (int j = startText; j <= endText; j++)
            //    {
            //        String tmpText = paragraph.GetRArray(i).GetTArray(j).StringValue;
            //        int startChar = 0, endChar = tmpText.Length() - 1;
            //        if ((j == textBegin) && (i == RunBegin))
            //            startChar = charBegin;
            //        if ((j == textEnd) && (i == RunEnd))
            //        {
            //            endChar = charEnd;
            //        }
            //        out1.Append(tmpText.Substring(startChar, endChar + 1));

            //    }
            //}
            //return out1.ToString();
            throw new NotImplementedException();
        }

        /**
         * Removes a Run at the position pos in the paragraph
         * @param pos
         * @return true if the run was Removed
         */
        public bool RemoveRun(int pos)
        {
            if (pos >= 0 && pos < paragraph.SizeOfRArray())
            {
                GetCTP().RemoveR(pos);
                runs.RemoveAt(pos);
                return true;
            }
            return false;
        }

        /**
         * returns the type of the BodyElement Paragraph
         * @see NPOI.XWPF.UserModel.IBodyElement#getElementType()
         */
        public BodyElementType ElementType
        {
            get
            {
                return BodyElementType.PARAGRAPH;
            }
        }

        public IBody Body
        {
            get
            {
                return part;
            }
        }

        /**
         * returns the part of the bodyElement
         * @see NPOI.XWPF.UserModel.IBody#getPart()
         */
        public POIXMLDocumentPart GetPart()
        {
            if (part != null)
            {
                return part.GetPart();
            }
            return null;
        }

        /**
         * returns the partType of the bodyPart which owns the bodyElement
         * 
         * @see NPOI.XWPF.UserModel.IBody#getPartType()
         */
        public BodyType PartType
        {
            get
            {
                return part.PartType;
            }
        }

        /**
         * Adds a new Run to the Paragraph
         * 
         * @param r
         */
        public void AddRun(XWPFRun r)
        {
            if (!runs.Contains(r))
            {
                runs.Add(r);
            }
        }

        /**
         * return the XWPFRun-Element which owns the CTR Run-Element
         * 
         * @param r
         */
        public XWPFRun GetRun(CT_R r)
        {
            for (int i = 0; i < Runs.Count; i++)
            {
                if (Runs[i].GetCTR() == r)
                {
                    return Runs[i];
                }
            }
            return null;
        }

    }//end class
}