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
    using NPOI.Util;
    using System.Collections;
    using NPOI.WP.UserModel;
    using S=NPOI.OpenXmlFormats.Shared;

    /**
* <p>A Paragraph within a Document, Table, Header etc.</p> 
* 
* <p>A paragraph has a lot of styling information, but the
*  actual text (possibly along with more styling) is held on
*  the child {@link XWPFRun}s.</p>
*/
    public class XWPFParagraph : IBodyElement, IRunBody, ISDTContents, IParagraph
    {
        private CT_P paragraph;
        protected IBody part;
        /** For access to the document's hyperlink, comments, tables etc */
        protected XWPFDocument document;
        protected List<XWPFRun> runs;
        protected List<IRunElement> iRuns;

        protected List<XWPFOMath> oMaths;

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
            iRuns = new List<IRunElement>();

            BuildRunsInOrderFromXml(paragraph.Items);

            oMaths = new List<XWPFOMath>();
            BuildOMathsInOrderFromXml(paragraph.Items);
            // Look for bits associated with the runs
            foreach (XWPFRun run in runs)
            {
                CT_R r = run.GetCTR();
                if (document != null)
                {
                    for (int i = 0; i < r.Items.Count; i++)
                    {
                        object o = r.Items[i];
                        if (o is CT_FtnEdnRef)
                        {
                            CT_FtnEdnRef ftn = (CT_FtnEdnRef)o;
                            footnoteText.Append("[").Append(ftn.id).Append(": ");

                            XWPFFootnote footnote = null;

                            if (r.ItemsElementName.Count > i && r.ItemsElementName[i] == RunItemsChoiceType.endnoteReference)
                            {
                                footnote = document.GetEndnoteByID(int.Parse(ftn.id));
                                if (footnote == null)
                                    footnote = document.GetFootnoteByID(int.Parse(ftn.id));
                            }
                            else
                            {
                                footnote = document.GetFootnoteByID(int.Parse(ftn.id));
                                if (footnote == null)
                                    footnote = document.GetEndnoteByID(int.Parse(ftn.id));
                            }

                            if (footnote != null)
                            {
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
                            }

                            footnoteText.Append("]");
                        }
                    }
                }
            }
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
                    XWPFRun r = new XWPFRun((CT_R)o, (IRunBody)this);
                    runs.Add(r);
                    iRuns.Add(r);
                }
                if (o is CT_Hyperlink1)
                {
                    CT_Hyperlink1 link = (CT_Hyperlink1)o;
                    foreach (CT_R r in link.GetRList())
                    {
                        //runs.Add(new XWPFHyperlinkRun(link, r, this));
                        XWPFHyperlinkRun hr = new XWPFHyperlinkRun(link, r, this);
                        runs.Add(hr);
                        iRuns.Add(hr);

                    }
                }
                if (o is CT_SimpleField) {
                    CT_SimpleField field = (CT_SimpleField)o;
                    foreach (CT_R r in field.GetRList())
                    {
                        XWPFFieldRun fr = new XWPFFieldRun(field, r, this);
                        runs.Add(fr);
                        iRuns.Add(fr);
                    }
                }
                if (o is CT_SdtBlock)
                {
                    XWPFSDT cc = new XWPFSDT((CT_SdtBlock)o, part);
                    iRuns.Add(cc);
                }
                if (o is CT_SdtRun)
                {
                    XWPFSDT cc = new XWPFSDT((CT_SdtRun)o, part);
                    iRuns.Add(cc);
                }
                if (o is CT_RunTrackChange)
                {
                    foreach (CT_R r in ((CT_RunTrackChange)o).GetRList())
                    {
                        XWPFRun cr = new XWPFRun(r, (IRunBody)this);
                        runs.Add(cr);
                        iRuns.Add(cr);
                    }
                }
                if (o is CT_SmartTagRun)
                {
                    // Smart Tags can be nested many times. 
                    // This implementation does not preserve the tagging information
                    BuildRunsInOrderFromXml((o as CT_SmartTagRun).Items);
                }
                if (o is CT_RunTrackChange) {
                    // add all the insertions as text
                    foreach (CT_RunTrackChange ins in ((CT_RunTrackChange)o).GetInsList())
                    {
                        foreach (CT_R r in ins.GetRList())
                        {
                            XWPFRun cr = new XWPFRun(r, (IRunBody)this);
                            runs.Add(cr);
                            iRuns.Add(cr);
                        }
                    }
                }
            }
        }

        private void BuildOMathsInOrderFromXml(ArrayList items)
        {
            foreach (object o in items)
            {
                if(o is S.CT_OMath)
                {
                    oMaths.Add(new XWPFOMath(o as S.CT_OMath, this));
                }
            }
        }


        public CT_P GetCTP()
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

        /**
         * Return literal runs and sdt/content control objects.
         * @return List<IRunElement>
         */
        public List<IRunElement> IRuns
        {
            get
            {
                return iRuns;
            }
        }

        public IList<XWPFOMath> OMaths
        {
            get
            {
                return oMaths.AsReadOnly();
            }
        }

        public bool IsEmpty
        {
            get
            {
                //!paragraph.getDomNode().hasChildNodes();
                //inner xml include objects holded by Items and pPr object
                //should use children of pPr node, but we didn't keep reference to it.
                return paragraph.Items.Count == 0 && (paragraph.pPr == null || paragraph.pPr.IsEmpty);
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
         * and std element in it.
         */
        public String Text
        {
            get
            {
                StringBuilder out1 = new StringBuilder();
                foreach (IRunElement run in iRuns)
                {
                    if (run is XWPFRun)
                    {
                        XWPFRun xRun = (XWPFRun)run;
                        // don't include the text if reviewing is enabled and this is a deleted run
                        if (xRun.GetCTR().GetDelTextList().Count==0)
                        {
                            out1.Append(xRun.ToString());
                        }
                    }
                    else if (run is XWPFSDT)
                    {
                        out1.Append(((XWPFSDT)run).Content.Text);
                    }
                    else
                    {
                        out1.Append(run.ToString());
                    }
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
         * Returns Ilvl of the numeric style for this paragraph.
         * Returns null if this paragraph does not have numeric style.
         * @return Ilvl as BigInteger
         */
        public string GetNumIlvl()
        {
            if (paragraph.pPr != null)
            {
                if (paragraph.pPr.numPr != null)
                {
                    if (paragraph.pPr.numPr.ilvl != null)
                        return paragraph.pPr.numPr.ilvl.val;
                }
            }
            return null;
        }

        /**
         * Returns numbering format for this paragraph, eg bullet or
         *  lowerLetter.
         * Returns null if this paragraph does not have numeric style.
         */
        public String GetNumFmt()
        {
            string numID = GetNumID();
            XWPFNumbering numbering = document.GetNumbering();
            if (numID != null && numbering != null)
            {
                XWPFNum num = numbering.GetNum(numID);
                if (num != null)
                {
                    string ilvl = GetNumIlvl();
                    string abstractNumId = num.GetCTNum().abstractNumId.val;
                    CT_AbstractNum anum = numbering.GetAbstractNum(abstractNumId).GetAbstractNum();
                    CT_Lvl level = null;
                    for (int i = 0; i < anum.lvl.Count; i++)
                    {
                        CT_Lvl lvl = anum.lvl[i];
                        if (lvl.ilvl.Equals(ilvl))
                        {
                            level = lvl;
                            break;
                        }
                    }
                    if (level != null && level.numFmt != null)
                        return level.numFmt.val.ToString();
                }
            }
            return null;
        }

        /**
     * Returns the text that should be used around the paragraph level numbers.
     *
     * @return a string (e.g. "%1.") or null if the value is not found.
     */
        public String NumLevelText
        {
            get
            {
                string numID = GetNumID();
                XWPFNumbering numbering = document.CreateNumbering();
                if (numID != null && numbering != null)
                {
                    XWPFNum num = numbering.GetNum(numID);
                    if (num != null)
                    {
                        string ilvl = GetNumIlvl();
                        CT_Num ctNum = num.GetCTNum();
                        if (ctNum == null)
                            return null;

                        CT_DecimalNumber ctDecimalNumber = ctNum.abstractNumId;
                        if (ctDecimalNumber == null)
                            return null;

                        string abstractNumId = ctDecimalNumber.val;
                        if (abstractNumId == null)
                            return null;

                        XWPFAbstractNum xwpfAbstractNum = numbering.GetAbstractNum(abstractNumId);

                        if (xwpfAbstractNum == null)
                            return null;

                        CT_AbstractNum anum = xwpfAbstractNum.GetCTAbstractNum();

                        if (anum == null)
                            return null;

                        CT_Lvl level = null;
                        for (int i = 0; i < anum.SizeOfLvlArray(); i++)
                        {
                            CT_Lvl lvl = anum.GetLvlArray(i);
                            if (lvl != null && lvl.ilvl != null && lvl.ilvl.Equals(ilvl))
                            {
                                level = lvl;
                                break;
                            }
                        }
                        if (level != null && level.lvlText != null
                            && level.lvlText.val != null)
                            return level.lvlText.val.ToString();
                    }
                }
                return null;
            }
        }


        /**
         * Gets the numstartOverride for the paragraph numbering for this paragraph.
         * @return returns the overridden start number or null if there is no override for this paragraph.
         */
        public string GetNumStartOverride()
        {
            string numID = GetNumID();
            XWPFNumbering numbering = document.CreateNumbering();
            if (numID != null && numbering != null)
            {
                XWPFNum num = numbering.GetNum(numID);

                if (num != null)
                {
                    CT_Num ctNum = num.GetCTNum();
                    if (ctNum == null)
                    {
                        return null;
                    }
                    string ilvl = GetNumIlvl();
                    CT_NumLvl level = null;
                    for (int i = 0; i < ctNum.SizeOfLvlOverrideArray(); i++)
                    {
                        CT_NumLvl ctNumLvl = ctNum.GetLvlOverrideArray(i);
                        if (ctNumLvl != null && ctNumLvl.ilvl != null &&
                            ctNumLvl.ilvl.Equals(ilvl))
                        {
                            level = ctNumLvl;
                            break;
                        }
                    }
                    if (level != null && level.startOverride != null)
                    {
                        return level.startOverride.val;
                    }
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
                StringBuilder text = new StringBuilder();
                foreach (XWPFRun run in runs)
                {
                    text.Append(run.ToString());
                }
                return text.ToString();
            }
        }

        /**
         * Returns any text from any suitable pictures in the paragraph
         */
        public String PictureText
        {
            get
            {
                StringBuilder text = new StringBuilder();
                foreach (XWPFRun run in runs)
                {
                    text.Append(run.PictureText);
                }
                return text.ToString();
            }
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
        public ParagraphAlignment Alignment
        {
            get
            {
                CT_PPr pr = GetCTPPr();
                return pr == null || !pr.IsSetJc() ? ParagraphAlignment.LEFT : EnumConverter.ValueOf<ParagraphAlignment, ST_Jc>(pr.jc.val);
            }
            set
            {
                CT_PPr pr = GetCTPPr();
                CT_Jc jc = pr.IsSetJc() ? pr.jc : pr.AddNewJc();
                jc.val = EnumConverter.ValueOf<ST_Jc, ParagraphAlignment>(value);
            }
        }

        
        /**
         * @return The raw alignment value, {@link #getAlignment()} is suggested
         */

        public int FontAlignment
        {
            get
            {
                return (int)Alignment;
            }
            set
            {
                Alignment = (ParagraphAlignment)value;
            }
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
        public TextAlignment VerticalAlignment
        {
            get
            {
                CT_PPr pr = GetCTPPr();
                return (pr == null || !pr.IsSetTextAlignment()) ? TextAlignment.AUTO
                        : EnumConverter.ValueOf<TextAlignment, ST_TextAlignment>(pr.textAlignment.val);
            }
            set
            {
                CT_PPr pr = GetCTPPr();
                CT_TextAlignment textAlignment = pr.IsSetTextAlignment() ? pr
                        .textAlignment : pr.AddNewTextAlignment();
                //STTextAlignment.Enum en = STTextAlignment.Enum
                //        .forInt(valign.Value);
                textAlignment.val = EnumConverter.ValueOf<ST_TextAlignment, TextAlignment>(value);
            }
        }

        /// <summary>
        /// the top border for the paragraph
        /// </summary>
        public Borders BorderTop
        {
            get
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
            set
            {
                CT_PBdr ct = GetCTPBrd(true);
                if (ct == null)
                {
                    throw new RuntimeException("invalid paragraph state");
                }
                CT_Border pr = ct.IsSetTop() ? ct.top : ct.AddNewTop();
                if (value == Borders.None)
                    ct.UnsetTop();
                else
                    pr.val = EnumConverter.ValueOf<ST_Border, Borders>(value);
            }
        }



        /// <summary>
        ///Specifies the border which shall be displayed below a Set of
        /// paragraphs which have the same Set of paragraph border Settings.
        /// </summary>
        /// <returns>the bottom border for the paragraph</returns>
        public Borders BorderBottom
        {
            get
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
            set
            {
                CT_PBdr ct = GetCTPBrd(true);
                CT_Border pr = ct.IsSetBottom() ? ct.bottom : ct.AddNewBottom();
                if (value == Borders.None)
                    ct.UnsetBottom();
                else
                    pr.val = EnumConverter.ValueOf<ST_Border, Borders>(value);
            }

        }

        /// <summary>
        /// Specifies the border which shall be displayed on the left side of the
        /// page around the specified paragraph.
        /// </summary>
        /// <returns>the left border for the paragraph</returns>
        public Borders BorderLeft
        {
            get
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
            set
            {
                CT_PBdr ct = GetCTPBrd(true);
                CT_Border pr = ct.IsSetLeft() ? ct.left : ct.AddNewLeft();
                if (value == Borders.None)
                    ct.UnsetLeft();
                else
                    pr.val = EnumConverter.ValueOf<ST_Border, Borders>(value);
            }
        }


        /**
         * Specifies the border which shall be displayed on the right side of the
         * page around the specified paragraph.
         *
         * @return ParagraphBorder - the right border for the paragraph
         * @see #setBorderRight(Borders)
         * @see Borders for a list of all possible borders
         */
        public Borders BorderRight
        {
            get
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
            set
            {
                CT_PBdr ct = GetCTPBrd(true);
                CT_Border pr = ct.IsSetRight() ? ct.right : ct.AddNewRight();
                if (value == Borders.None)
                    ct.UnsetRight();
                else
                    pr.val = EnumConverter.ValueOf<ST_Border, Borders>(value);
            }
        }
        public ST_Shd FillPattern
        {
            get
            {
                if (!this.GetCTPPr().IsSetShd())
                    return ST_Shd.nil;

                return this.GetCTPPr().shd.val;
            }
            set
            {
                CT_Shd ctShd = null;
                if (!this.GetCTPPr().IsSetShd())
                {
                    ctShd = this.GetCTPPr().AddNewShd();
                }
                else
                {
                    ctShd = this.GetCTPPr().shd;
                }
                ctShd.val = value;
            }
        }
        public string FillBackgroundColor
        {
            get
            {
                if (!this.GetCTPPr().IsSetShd())
                    return null;

                return this.GetCTPPr().shd.fill;
            }
            set
            {
                CT_Shd ctShd = null;
                if (!this.GetCTPPr().IsSetShd())
                {
                    ctShd = this.GetCTPPr().AddNewShd();
                }
                else
                {
                    ctShd = this.GetCTPPr().shd;
                }
                ctShd.color = "auto";
                ctShd.fill = value;
            }
        }
        /**
         * Specifies the border which shall be displayed between each paragraph in a
         * Set of paragraphs which have the same Set of paragraph border Settings.
         *
         * @return ParagraphBorder - the between border for the paragraph
         * @see #setBorderBetween(Borders)
         * @see Borders for a list of all possible borders
         */
        public Borders BorderBetween
        {
            get
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
            set
            {
                CT_PBdr ct = GetCTPBrd(true);
                CT_Border pr = ct.IsSetBetween() ? ct.between : ct.AddNewBetween();
                if (value == Borders.None)
                    ct.UnsetBetween();
                else
                    pr.val = EnumConverter.ValueOf<ST_Border, Borders>(value);
            }
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
            set
            {
                CT_PPr ppr = GetCTPPr();
                CT_OnOff ct_pageBreak = ppr.IsSetPageBreakBefore() ? ppr
                        .pageBreakBefore : ppr.AddNewPageBreakBefore();
                ct_pageBreak.val = value;
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
         * paragraph in the document in absolute units.
         *
         * @return bigint - value representing the spacing After the paragraph
         * @see #setSpacingAfterLines(int)
         */
        public int SpacingAfterLines
        {
            get
            {
                CT_Spacing spacing = GetCTSpacing(false);
                return (spacing != null && spacing.IsSetAfterLines()) ? int.Parse(spacing.afterLines) : -1;
            }
            set
            {
                CT_Spacing spacing = GetCTSpacing(true);
                //BigInteger bi = new BigInteger("" + spaces);
                spacing.afterLines = value.ToString();
            }
        }

        /**
         * Specifies the spacing that should be Added above the first line in this
         * paragraph in the document in absolute units.
         *
         * @return the spacing that should be Added above the first line
         * @see #setSpacingBefore(int)
         */
        public int SpacingBefore
        {
            get
            {
                CT_Spacing spacing = GetCTSpacing(false);
                return (spacing != null && spacing.IsSetBefore()) ? (int)spacing.before : -1;
            }
            set
            {
                CT_Spacing spacing = GetCTSpacing(true);
                //BigInteger bi = new BigInteger("" + spaces);
                spacing.before = (ulong)value;
            }
        }


        /**
         * Specifies the spacing that should be Added before the first line in this paragraph in the
         * document in line units.
         * The value of this attribute is specified in one hundredths of a line.
         *
         * @return the spacing that should be Added before the first line in this paragraph
         * @see #setSpacingBeforeLines(int)
         */
        public int SpacingBeforeLines
        {
            get
            {
                CT_Spacing spacing = GetCTSpacing(false);
                return (spacing != null && spacing.IsSetBeforeLines()) ? int.Parse(spacing.beforeLines) : -1;
            }

            set
            {
                CT_Spacing spacing = GetCTSpacing(true);
                //BigInteger bi = new BigInteger("" + spaces);
                spacing.beforeLines = value.ToString();
            }
        }

        /// <summary>
        ///Specifies how the spacing between lines is calculated as stored in the
        /// line attribute. If this attribute is omitted, then it shall be assumed to
        /// be of a value auto if a line attribute value is present.
        /// </summary>
        public LineSpacingRule SpacingLineRule
        {
            get
            {
                CT_Spacing spacing = GetCTSpacing(false);
                return (spacing != null && spacing.IsSetLineRule()) ?
                    EnumConverter.ValueOf<LineSpacingRule, ST_LineSpacingRule>(spacing.lineRule) : LineSpacingRule.AUTO;
            }
            set
            {
                CT_Spacing spacing = GetCTSpacing(true);
                spacing.lineRule = EnumConverter.ValueOf<ST_LineSpacingRule, LineSpacingRule>(value);
            }
        }

        ///<summary>
        /// Return the spacing between lines of a paragraph. The units of the return value depends on the
        /// <see cref="LineSpacingRule"/>. If AUTO, the return value is in lines, otherwise the return
        /// value is in points
        /// 
        /// <return>a double specifying points or lines.</return>
        ///</summary>
        public double SpacingBetween
        {
            set
            {
                setSpacingBetween(value, LineSpacingRule.AUTO);
            }

        }
        public void setSpacingBetween(double spacing, LineSpacingRule rule)
        {
            CT_Spacing ctSp = GetCTSpacing(true);
            switch(rule)
            {
                case LineSpacingRule.AUTO:
                    ctSp.line = Math.Round(spacing * 240.0).ToString();
                    break;
                default:
                    ctSp.line = Math.Round(spacing * 20.0).ToString();
                    break;
            }
            ctSp.lineRule = EnumConverter.ValueOf<ST_LineSpacingRule, LineSpacingRule>(rule);

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

        public int IndentFromLeft
        {
            get
            {
                return IndentationLeft;
            }
            set
            {
                IndentationLeft = value;
            }
        }

        public int IndentFromRight
        {
            get
            {
                return IndentationRight;
            }
            set
            {
                IndentationRight = value;
            }
        }

        public int FirstLineIndent
        {
            get
            {
                return IndentationFirstLine;
            }
            set
            {
                IndentationFirstLine = (value);
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
        public bool IsWordWrapped
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
        [Obsolete]
        public bool IsWordWrap
        {
            get { return IsWordWrapped; }
            set { IsWordWrapped = value; }
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
        internal CT_PPr GetCTPPr()
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
        protected internal void AddRun(CT_R run)
        {
            int pos= paragraph.GetRList().Count;
            paragraph.AddNewR();
            paragraph.SetRArray(pos, run);
        }
        /// <summary>
        /// Replace text inside each run (cross run is not supported yet)
        /// </summary>
        /// <param name="oldText">target text</param>
        /// <param name="newText">replacement text</param>
        public void ReplaceText(string oldText, string newText)
        {
            TextSegment ts= this.SearchText(oldText, new PositionInParagraph() { Run = 0 });
            if (ts.BeginRun == ts.EndRun)
            {
                this.runs[ts.BeginRun].ReplaceText(oldText, newText);
            }
            else
            {
                this.runs[ts.BeginRun].ReplaceText(this.runs[ts.BeginRun].Text.Substring(ts.BeginChar), newText);
                this.runs[ts.EndRun].ReplaceText(this.runs[ts.EndRun].Text.Substring(0, ts.EndChar + 1), "");
                for (int i = ts.EndRun-1; i > ts.BeginRun; i--)
                {
                    RemoveRun(i);
                }
            }
        }
        /// <summary>
        /// this methods parse the paragraph and search for the string searched. 
        /// If it finds the string, it will return true and the position of the String will be saved in the parameter startPos.
        /// </summary>
        /// <param name="searched"></param>
        /// <param name="startPos"></param>
        /// <returns></returns>
        public TextSegment SearchText(String searched, PositionInParagraph startPos)
        {

            int startRun = startPos.Run,
                startText = startPos.Text,
                startChar = startPos.Char;
            int beginRunPos = 0, beginTextPos = 0, beginCharPos = 0,candCharPos = 0;
            bool newList = false;
            for (int runPos = startRun; runPos < paragraph.GetRList().Count; runPos++)
            {
                int  textPos = 0, charPos = 0;
                CT_R ctRun = paragraph.GetRList()[runPos];
                foreach (object o in ctRun.Items)
                {
                    if (o is CT_Text)
                    {
                        if (textPos >= startText)
                        {
                            String candidate = ((CT_Text)o).Value;
                            if (runPos == startRun)
                                charPos = startChar;
                            else
                                charPos = 0;
                            for (; charPos < candidate.Length; charPos++)
                            {
                                if ((candidate[charPos] == searched[0]) && (candCharPos == 0))
                                {
                                    beginTextPos = textPos;
                                    beginCharPos = charPos;
                                    beginRunPos = runPos;
                                    newList = true;
                                }
                                if (candidate[charPos] == searched[candCharPos])
                                {
                                    if (candCharPos + 1 < searched.Length)
                                    {
                                        candCharPos++;
                                    }
                                    else if (newList)
                                    {
                                        TextSegment segement = new TextSegment();
                                        segement.BeginRun = (beginRunPos);
                                        segement.BeginText = (beginTextPos);
                                        segement.BeginChar = (beginCharPos);
                                        segement.EndRun = (runPos);
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
         * Appends a new run to this paragraph
         *
         * @return a new text run
         */
        public XWPFRun CreateRun()
        {
            XWPFRun xwpfRun = new XWPFRun(paragraph.AddNewR(), (IRunBody)this);
            runs.Add(xwpfRun);
            iRuns.Add(xwpfRun);
            return xwpfRun;
        }

        /**
         * Appends a new OMath to this paragraph
         *
         * @return a new text run
         */
        public XWPFOMath CreateOMath()
        {
            XWPFOMath oMath = new XWPFOMath(paragraph.AddNewOMath(), this);
            oMaths.Add(oMath);            
            return oMath;
        }

        /// <summary>
        /// insert a new Run in RunArray
        /// </summary>
        /// <param name="pos">The position at which the new run should be added.</param>
        /// <returns>the inserted run or null if the given pos is out of bounds.</returns>
        public XWPFRun InsertNewRun(int pos)
        {
            if (pos >= 0 && pos <= runs.Count)
            {
                // calculate the correct pos as our run/irun list contains
                // hyperlinks
                // and fields so it is different to the paragraph R array.
                int rPos = 0;
                for (int i = 0; i < pos; i++)
                {
                    XWPFRun currRun = runs[i];
                    if (!(currRun is XWPFHyperlinkRun
                        || currRun is XWPFFieldRun))
                    {
                        rPos++;
                    }
                }

            CT_R ctRun = paragraph.InsertNewR(rPos);
            XWPFRun newRun = new XWPFRun(ctRun, (IRunBody)this);

            // To update the iRuns, find where we're going
            // in the normal Runs, and go in there
            int iPos = iRuns.Count;
                if (pos < runs.Count)
                {
                    XWPFRun oldAtPos = runs[(pos)];
                    int oldAt = iRuns.IndexOf(oldAtPos);
                    if (oldAt != -1)
                    {
                        iPos = oldAt;
                    }
                }
                iRuns.Insert(iPos, newRun);

                // Runs itself is easy to update
                runs.Insert(pos, newRun);

                return newRun;
            }
            return null;
        }


        /**
         * Get a Text
         * @param segment
         */
        public String GetText(TextSegment segment)
        {
            int RunBegin = segment.BeginRun;
            int textBegin = segment.BeginText;
            int charBegin = segment.BeginChar;
            int RunEnd = segment.EndRun;
            int textEnd = segment.EndText;
            int charEnd = segment.EndChar;
            StringBuilder text = new StringBuilder();
            for (int i = RunBegin; i <= RunEnd; i++)
            {
                int startText = 0, endText = paragraph.GetRList()[i].GetTList().Count - 1;
                if (i == RunBegin)
                    startText = textBegin;
                if (i == RunEnd)
                    endText = textEnd;
                for (int j = startText; j <= endText; j++)
                {
                    String tmpText = paragraph.GetRList()[i].GetTArray(j).Value;
                    int startChar = 0, endChar = tmpText.Length - 1;
                    if ((j == textBegin) && (i == RunBegin))
                        startChar = charBegin;
                    if ((j == textEnd) && (i == RunEnd))
                    {
                        endChar = charEnd;
                    }
                    text.Append(tmpText.Substring(startChar, endChar - startChar + 1));

                }
            }
            return text.ToString();
        }

        /**
         * Removes a Run at the position pos in the paragraph
         * @param pos
         * @return true if the run was Removed
         */
        public bool RemoveRun(int pos)
        {
            if (pos >= 0 && pos < runs.Count)
            {
                // Remove the run from our high level lists
                XWPFRun run = runs[(pos)];
                if (run is XWPFHyperlinkRun || run is XWPFFieldRun)
                {
                    // TODO Add support for removing these kinds of nested runs,
                    //  which aren't on the CTP -> R array, but CTP -> XXX -> R array
                    throw new ArgumentException("Removing Field or Hyperlink runs not yet supported");
                }
                runs.RemoveAt(pos);
                iRuns.Remove(run);
                // Remove the run from the low-level XML
                //calculate the correct pos as our run/irun list contains hyperlinks and fields so is different to the paragraph R array.
                int rPos = 0;
                for (int i = 0; i < pos; i++)
                {
                    XWPFRun currRun = runs[i];
                    if (!(currRun is XWPFHyperlinkRun || currRun is XWPFFieldRun))
                        rPos++;
                }
                GetCTP().RemoveR(pos);
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
        public POIXMLDocumentPart Part
        {
            get
            {
                if (part != null)
                {
                    return part.Part;
                }
                return null;
            }
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
            for (int i = 0; i < runs.Count; i++)
            {
                if (runs[i].GetCTR() == r)
                {
                    return runs[i];
                }
            }
            return null;
        }
        /// <summary>
        /// Appends a new hyperlink run to this paragraph
        /// </summary>
        /// <param name="rId">a new hyperlink run</param>
        /// <returns></returns>
        public XWPFHyperlinkRun CreateHyperlinkRun(string rId)
        {
            CT_R r = new CT_R();
            r.AddNewRPr().rStyle = new CT_String() { val = "Hyperlink" };

            CT_Hyperlink1 hl = paragraph.AddNewHyperlink();
            hl.history = ST_OnOff.on;
            hl.id = rId;
            hl.Items.Add(r);
            XWPFHyperlinkRun xwpfRun = new XWPFHyperlinkRun(hl, r, this);
            runs.Add(xwpfRun);
            iRuns.Add(xwpfRun);
            return xwpfRun;
        }
    }

}