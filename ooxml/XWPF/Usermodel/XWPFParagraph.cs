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
    /**
     * Sketch of XWPF paragraph class
     */
    public class XWPFParagraph : IBodyElement
    {
        private CT_P paragraph;
        protected IBody part;
        /** For access to the document's hyperlink, comments, tables etc */
        protected XWPFDocument document;
        protected List<XWPFRun> Runs;

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

            Runs = new List<XWPFRun>();

            foreach (object o in paragraph.Items)
            {
                if (o is CT_R)
                {
                    Runs.Add(new XWPFRun((CT_R)o, this));
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
            }

            c.Dispose();
       
            // Look for bits associated with the Runs
            foreach(XWPFRun run in Runs) {
               CTR r = Run.CTR;
          
               // Check for bits that only apply when
               //  attached to a core document
               // TODO Make this nicer by tracking the XWPFFootnotes directly
               if(document != null) {
                  c = r.NewCursor();
                  c.SelectPath("child::*");
                  while (c.ToNextSelection()) {
                     XmlObject o = c.Object;
                     if(o is CTFtnEdnRef) {
                        CTFtnEdnRef ftn = (CTFtnEdnRef)o;
                        footnoteText.Append("[").Append(ftn.Id).Append(": ");
                        XWPFFootnote footnote =
                           ftn.DomNode.LocalName.Equals("footnoteReference") ?
                                 document.GetFootnoteByID(ftn.Id.IntValue()) :
                                 document.GetEndnoteByID(ftn.Id.IntValue());
   
                        bool first = true;
                        foreach (XWPFParagraph p in footnote.Paragraphs) {
                           if (!first) {
                              footnoteText.Append("\n");
                              first = false;
                           }
                           footnoteText.Append(p.Text);
                        }
   
                        footnoteText.Append("]");
                     }
                  }
                  c.Dispose();
               }
           }*/
        }


        public CT_P GetCTP()
        {
            return paragraph;
        }

        public List<XWPFRun> GetRuns()
        {
            //return Collections.UnmodifiableList(Runs);
            throw new NotImplementedException();
        }

        public bool IsEmpty()
        {
            //return !paragraph.DomNode.HasChildNodes();
            throw new NotImplementedException();
        }

        public XWPFDocument GetDocument()
        {
            return document;
        }

        /**
         * Return the textual content of the paragraph, including text from pictures
         * in it.
         */
        public String GetText()
        {
            //StringBuilder out1 = new StringBuilder();
            //foreach(XWPFRun run in Runs) {
            //   out1.Append(Run.ToString());
            //}
            //out1.Append(footnoteText);
            //return out1.ToString();
            throw new NotImplementedException();
        }

        /**
         * Return styleID of the paragraph if style exist for this paragraph
         * if not, null will be returned     
         * @return		styleID as String
         */
        public String GetStyleID()
        {
            //if (paragraph.PPr != null){
            //    if(paragraph.PPr.PStyle!= null){
            //        if (paragraph.PPr.PStyle.Val!= null)
            //            return paragraph.PPr.PStyle.Val;
            //    }
            //}
            //return null;
            throw new NotImplementedException();
        }
        /**
         * If style exist for this paragraph
         * NumId of the paragraph will be returned.
         * If style not exist null will be returned     
         * @return	NumID as Bigint
         */
        //public Bigint GetNumID(){
        //    if(paragraph.PPr!=null){
        //        if(paragraph.PPr.NumPr!=null){
        //            if(paragraph.PPr.NumPr.NumId!=null)
        //                return paragraph.PPr.NumPr.NumId.Val;
        //        }
        //    }
        //    return null;
        //}

        /**
         * SetNumID of Paragraph
         * @param numPos
         */
        //public void SetNumID(Bigint numPos) {
        //    if(paragraph.PPr==null)
        //        paragraph.AddNewPPr();
        //    if(paragraph.PPr.NumPr==null)
        //        paragraph.PPr.AddNewNumPr();
        //    if(paragraph.PPr.NumPr.NumId==null){
        //        paragraph.PPr.NumPr.AddNewNumId();
        //    }
        //    paragraph.PPr.NumPr.NumId.Val=(numPos);
        //}

        /**
         * Returns the text of the paragraph, but not of any objects in the
         * paragraph
         */
        public String GetParagraphText()
        {
            //StringBuilder out1 = new StringBuilder();
            //foreach(XWPFRun run in Runs) {
            //   out1.Append(Run.ToString());
            //}
            //return out1.ToString();
            throw new NotImplementedException();
        }

        /**
         * Returns any text from any suitable pictures in the paragraph
         */
        public String GetPictureText()
        {
            //StringBuilder out1 = new StringBuilder();
            //foreach(XWPFRun run in Runs) {
            //   out1.Append(Run.PictureText);
            //}
            //return out1.ToString();
            throw new NotImplementedException();
        }

        /**
         * Returns the footnote text of the paragraph
         *
         * @return  the footnote text or empty string if the paragraph does not have footnotes
         */
        public String GetFootnoteText()
        {
            return footnoteText.ToString();
        }

        /**
         * Appends a new run to this paragraph
         *
         * @return a new text run
         */
        public XWPFRun CreateRun()
        {
            XWPFRun xwpfRun = new XWPFRun(paragraph.AddNewR(), this);
            Runs.Add(xwpfRun);
            return xwpfRun;
        }

        /**
         * Returns the paragraph alignment which shall be applied to text in this
         * paragraph.
         * <p/>
         * <p/>
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
            //CTPPr pr = GetCTPPr();
            //return pr == null || !pr.IsSetJc() ? ParagraphAlignment.LEFT
            //        : ParagraphAlignment.ValueOf(pr.Jc.Val.IntValue());
            throw new NotImplementedException();
        }

        /**
         * Specifies the paragraph alignment which shall be applied to text in this
         * paragraph.
         * <p/>
         * <p/>
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
            //CTPPr pr = GetCTPPr();
            //CTJc jc = pr.IsSetJc() ? pr.Jc : pr.AddNewJc();
            //STJc.Enum en = STJc.Enum.ForInt(align.Value);
            //jc.Val=(en);
            throw new NotImplementedException();
        }

        /**
         * Returns the text vertical alignment which shall be applied to text in
         * this paragraph.
         * <p/>
         * If the line height (before any Added spacing) is larger than one or more
         * characters on the line, all characters will be aligned to each other as
         * specified by this element.
         * </p>
         * <p/>
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
            //CTPPr pr = GetCTPPr();
            //return (pr == null || !pr.IsSetTextAlignment()) ? TextAlignment.AUTO
            //        : TextAlignment.ValueOf(pr.TextAlignment.Val
            //        .intValue());
            throw new NotImplementedException();
        }

        /**
         * Specifies the text vertical alignment which shall be applied to text in
         * this paragraph.
         * <p/>
         * If the line height (before any Added spacing) is larger than one or more
         * characters on the line, all characters will be aligned to each other as
         * specified by this element.
         * </p>
         * <p/>
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
            //CTPPr pr = GetCTPPr();
            //CTTextAlignment textAlignment = pr.IsSetTextAlignment() ? pr
            //        .TextAlignment : pr.AddNewTextAlignment();
            //STTextAlignment.Enum en = STTextAlignment.Enum
            //        .forInt(valign.Value);
            //textAlignment.Val=(en);
            throw new NotImplementedException();
        }

        /**
         * Specifies the border which shall be displayed above a Set of paragraphs
         * which have the same Set of paragraph border Settings.
         * <p/>
         * <p/>
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
         * <p/>
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
            //CTPBdr ct = GetCTPBrd(true);

            //CTBorder pr = (ct != null && ct.IsSetTop()) ? ct.Top : ct.AddNewTop();
            //if (border.Value == Borders.NONE.Value)
            //    ct.UnsetTop();
            //else
            //    pr.Val=(STBorder.Enum.ForInt(border.Value));
            throw new NotImplementedException();
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
            //CTPBdr border = GetCTPBrd(false);
            //CTBorder ct = null;
            //if (border != null) {
            //    ct = border.Top;
            //}
            //STBorder.Enum ptrn = (ct != null) ? ct.Val : STBorder.NONE;
            //return Borders.ValueOf(ptrn.IntValue());
            throw new NotImplementedException();
        }

        /**
         * Specifies the border which shall be displayed below a Set of paragraphs
         * which have the same Set of paragraph border Settings.
         * <p/>
         * To determine if any two adjoining paragraphs shall have an individual top
         * and bottom border or a between border, the Set of borders on the two
         * adjoining paragraphs are Compared. If the border information on those two
         * paragraphs is identical for all possible paragraphs borders, then the
         * between border is displayed. Otherwise, the paragraph shall use its
         * bottom border and the following paragraph shall use its top border,
         * respectively. If this border specifies a space attribute, that value
         * determines the space After the bottom of the text (ignoring any space
         * below) which should be left before this border is Drawn, specified in
         * points.
         * </p>
         * <p/>
         * If this element is omitted on a given paragraph, its value is determined
         * by the Setting previously Set at any level of the style hierarchy (i.e.
         * that previous Setting remains unChanged). If this Setting is never
         * specified in the style hierarchy, then no between border shall be applied
         * below identical paragraphs.
         * </p>
         * <b>This border can only be a line border.</b>
         *
         * @param border
         * @see Borders a list of all types of borders
         */
        public void SetBorderBottom(Borders border)
        {
            //CTPBdr ct = GetCTPBrd(true);
            //CTBorder pr = ct.IsSetBottom() ? ct.Bottom : ct.AddNewBottom();
            //if (border.Value == Borders.NONE.Value)
            //    ct.UnsetBottom();
            //else
            //    pr.Val=(STBorder.Enum.ForInt(border.Value));
            throw new NotImplementedException();
        }

        /**
         * Specifies the border which shall be displayed below a Set of
         * paragraphs which have the same Set of paragraph border Settings.
         *
         * @return paragraphBorder - the bottom border for the paragraph
         * @see #setBorderBottom(Borders)
         * @see Borders a list of all types of borders
         */
        public Borders GetBorderBottom()
        {
            //CTPBdr border = GetCTPBrd(false);
            //CTBorder ct = null;
            //if (border != null) {
            //    ct = border.Bottom;
            //}
            //STBorder.Enum ptrn = ct != null ? ct.Val : STBorder.NONE;
            //return Borders.ValueOf(ptrn.IntValue());
            throw new NotImplementedException();
        }

        /**
         * Specifies the border which shall be displayed on the left side of the
         * page around the specified paragraph.
         * <p/>
         * To determine if any two adjoining paragraphs should have a left border
         * which spans the full line height or not, the left border shall be Drawn
         * between the top border or between border at the top (whichever would be
         * rendered for the current paragraph), and the bottom border or between
         * border at the bottom (whichever would be rendered for the current
         * paragraph).
         * </p>
         * <p/>
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
            //CTPBdr ct = GetCTPBrd(true);
            //CTBorder pr = ct.IsSetLeft() ? ct.Left : ct.AddNewLeft();
            //if (border.Value == Borders.NONE.Value)
            //    ct.UnsetLeft();
            //else
            //    pr.Val=(STBorder.Enum.ForInt(border.Value));
            throw new NotImplementedException();
        }

        /**
         * Specifies the border which shall be displayed on the left side of the
         * page around the specified paragraph.
         *
         * @return ParagraphBorder - the left border for the paragraph
         * @see #setBorderLeft(Borders)
         * @see Borders for a list of all possible borders
         */
        public Borders GetBorderLeft()
        {
            //CTPBdr border = GetCTPBrd(false);
            //CTBorder ct = null;
            //if (border != null) {
            //    ct = border.Left;
            //}
            //STBorder.Enum ptrn = ct != null ? ct.Val : STBorder.NONE;
            //return Borders.ValueOf(ptrn.IntValue());
            throw new NotImplementedException();
        }

        /**
         * Specifies the border which shall be displayed on the right side of the
         * page around the specified paragraph.
         * <p/>
         * To determine if any two adjoining paragraphs should have a right border
         * which spans the full line height or not, the right border shall be Drawn
         * between the top border or between border at the top (whichever would be
         * rendered for the current paragraph), and the bottom border or between
         * border at the bottom (whichever would be rendered for the current
         * paragraph).
         * </p>
         * <p/>
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
            //CTPBdr ct = GetCTPBrd(true);
            //CTBorder pr = ct.IsSetRight() ? ct.Right : ct.AddNewRight();
            //if (border.Value == Borders.NONE.Value)
            //    ct.UnsetRight();
            //else
            //    pr.Val=(STBorder.Enum.ForInt(border.Value));
            throw new NotImplementedException();
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
            //CTPBdr border = GetCTPBrd(false);
            //CTBorder ct = null;
            //if (border != null) {
            //    ct = border.Right;
            //}
            //STBorder.Enum ptrn = ct != null ? ct.Val : STBorder.NONE;
            //return Borders.ValueOf(ptrn.IntValue());
            throw new NotImplementedException();
        }

        /**
         * Specifies the border which shall be displayed between each paragraph in a
         * Set of paragraphs which have the same Set of paragraph border Settings.
         * <p/>
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
         * <p/>
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
            //CTPBdr ct = GetCTPBrd(true);
            //CTBorder pr = ct.IsSetBetween() ? ct.Between : ct.AddNewBetween();
            //if (border.Value == Borders.NONE.Value)
            //    ct.UnsetBetween();
            //else
            //    pr.Val=(STBorder.Enum.ForInt(border.Value));
            throw new NotImplementedException();
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
            //CTPBdr border = GetCTPBrd(false);
            //CTBorder ct = null;
            //if (border != null) {
            //    ct = border.Between;
            //}
            //STBorder.Enum ptrn = ct != null ? ct.Val : STBorder.NONE;
            //return Borders.ValueOf(ptrn.IntValue());
            throw new NotImplementedException();
        }

        /**
         * Specifies that when rendering this document in a paginated
         * view, the contents of this paragraph are rendered on the start of a new
         * page in the document.
         * <p/>
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
            //CTPPr ppr = GetCTPPr();
            //CTOnOff ct_pageBreak = ppr.IsSetPageBreakBefore() ? ppr
            //        .PageBreakBefore : ppr.AddNewPageBreakBefore();
            //if (pageBreak)
            //    ct_pageBreak.Val=(STOnOff.TRUE);
            //else
            //    ct_pageBreak.Val=(STOnOff.FALSE);
            throw new NotImplementedException();
        }

        /**
         * Specifies that when rendering this document in a paginated
         * view, the contents of this paragraph are rendered on the start of a new
         * page in the document.
         * <p/>
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
        public bool IsPageBreak()
        {
            //CTPPr ppr = GetCTPPr();
            //CTOnOff ct_pageBreak = ppr.IsSetPageBreakBefore() ? ppr
            //        .PageBreakBefore : null;
            //if (ct_pageBreak != null
            //        && ct_pageBreak.Val.IntValue() == STOnOff.INT_TRUE) {
            //    return true;
            //}
            //return false;
            throw new NotImplementedException();
        }

        /**
         * Specifies the spacing that should be Added After the last line in this
         * paragraph in the document in absolute units.
         * <p/>
         * If the AfterLines attribute or the AfterAutoSpacing attribute is also
         * specified, then this attribute value is ignored.
         * </p>
         *
         * @param spaces -
         *               a positive whole number, whose contents consist of a
         *               measurement in twentieths of a point.
         */
        public void SetSpacingAfter(int spaces)
        {
            //CTSpacing spacing = GetCTSpacing(true);
            //if (spacing != null) {
            //    Bigint bi = new Bigint("" + spaces);
            //    spacing.After=(bi);
            //}
            throw new NotImplementedException();
        }

        /**
         * Specifies the spacing that should be Added After the last line in this
         * paragraph in the document in absolute units.
         *
         * @return int - value representing the spacing After the paragraph
         */
        public int GetSpacingAfter()
        {
            //CTSpacing spacing = GetCTSpacing(false);
            //return (spacing != null && spacing.IsSetAfter()) ? spacing.After.IntValue() : -1;
            throw new NotImplementedException();
        }

        /**
         * Specifies the spacing that should be Added After the last line in this
         * paragraph in the document in line units.
         * <b>The value of this attribute is
         * specified in one hundredths of a line.
         * </b>
         * <p/>
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
            //CTSpacing spacing = GetCTSpacing(true);
            //Bigint bi = new Bigint("" + spaces);
            //spacing.AfterLines=(bi);
            throw new NotImplementedException();
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
            //CTSpacing spacing = GetCTSpacing(false);
            //return (spacing != null && spacing.IsSetAfterLines()) ? spacing.AfterLines.IntValue() : -1;
            throw new NotImplementedException();
        }


        /**
         * Specifies the spacing that should be Added above the first line in this
         * paragraph in the document in absolute units.
         * <p/>
         * If the beforeLines attribute or the beforeAutoSpacing attribute is also
         * specified, then this attribute value is ignored.
         * </p>
         *
         * @param spaces
         */
        public void SetSpacingBefore(int spaces)
        {
            //CTSpacing spacing = GetCTSpacing(true);
            //Bigint bi = new Bigint("" + spaces);
            //spacing.Before=(bi);
            throw new NotImplementedException();
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
            //CTSpacing spacing = GetCTSpacing(false);
            //return (spacing != null && spacing.IsSetBefore()) ? spacing.Before.IntValue() : -1;
            throw new NotImplementedException();
        }

        /**
         * Specifies the spacing that should be Added before the first line in this
         * paragraph in the document in line units. <b> The value of this attribute
         * is specified in one hundredths of a line. </b>
         * <p/>
         * If the beforeAutoSpacing attribute is also specified, then this attribute
         * value is ignored. If this Setting is never specified in the style
         * hierarchy, then its value shall be zero.
         * </p>
         *
         * @param spaces
         */
        public void SetSpacingBeforeLines(int spaces)
        {
            //CTSpacing spacing = GetCTSpacing(true);
            //Bigint bi = new Bigint("" + spaces);
            //spacing.BeforeLines=(bi);
            throw new NotImplementedException();
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
            //CTSpacing spacing = GetCTSpacing(false);
            //return (spacing != null && spacing.IsSetBeforeLines()) ? spacing.BeforeLines.IntValue() : -1;
            throw new NotImplementedException();
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
            //CTSpacing spacing = GetCTSpacing(true);
            //spacing.LineRule=(STLineSpacingRule.Enum.ForInt(rule.Value));
            throw new NotImplementedException();
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
            //CTSpacing spacing = GetCTSpacing(false);
            //return (spacing != null && spacing.IsSetLineRule()) ? LineSpacingRule.ValueOf(spacing
            //        .LineRule.IntValue()) : LineSpacingRule.AUTO;
            throw new NotImplementedException();
        }


        /**
         * Specifies the indentation which shall be placed between the left text
         * margin for this paragraph and the left edge of that paragraph's content
         * in a left to right paragraph, and the right text margin and the right
         * edge of that paragraph's text in a right to left paragraph
         * <p/>
         * If this attribute is omitted, its value shall be assumed to be zero.
         * Negative values are defined such that the text is Moved past the text margin,
         * positive values Move the text inside the text margin.
         * </p>
         *
         * @param indentation
         */
        public void SetIndentationLeft(int indentation)
        {
            //CTInd indent = GetCTInd(true);
            //Bigint bi = new Bigint("" + indentation);
            //indent.Left=(bi);
            throw new NotImplementedException();
        }

        /**
         * Specifies the indentation which shall be placed between the left text
         * margin for this paragraph and the left edge of that paragraph's content
         * in a left to right paragraph, and the right text margin and the right
         * edge of that paragraph's text in a right to left paragraph
         * <p/>
         * If this attribute is omitted, its value shall be assumed to be zero.
         * Negative values are defined such that the text is Moved past the text margin,
         * positive values Move the text inside the text margin.
         * </p>
         *
         * @return indentation or null if indentation is not Set
         */
        public int GetIndentationLeft()
        {
            //CTInd indentation = GetCTInd(false);
            //return (indentation != null && indentation.IsSetLeft()) ? indentation.Left.IntValue()
            //        : -1;
            throw new NotImplementedException();
        }

        /**
         * Specifies the indentation which shall be placed between the right text
         * margin for this paragraph and the right edge of that paragraph's content
         * in a left to right paragraph, and the right text margin and the right
         * edge of that paragraph's text in a right to left paragraph
         * <p/>
         * If this attribute is omitted, its value shall be assumed to be zero.
         * Negative values are defined such that the text is Moved past the text margin,
         * positive values Move the text inside the text margin.
         * </p>
         *
         * @param indentation
         */
        public void SetIndentationRight(int indentation)
        {
            //CTInd indent = GetCTInd(true);
            //Bigint bi = new Bigint("" + indentation);
            //indent.Right = (bi);
            throw new NotImplementedException();
        }

        /**
         * Specifies the indentation which shall be placed between the right text
         * margin for this paragraph and the right edge of that paragraph's content
         * in a left to right paragraph, and the right text margin and the right
         * edge of that paragraph's text in a right to left paragraph
         * <p/>
         * If this attribute is omitted, its value shall be assumed to be zero.
         * Negative values are defined such that the text is Moved past the text margin,
         * positive values Move the text inside the text margin.
         * </p>
         *
         * @return indentation or null if indentation is not Set
         */

        public int GetIndentationRight()
        {
            //CTInd indentation = GetCTInd(false);
            //return (indentation != null && indentation.IsSetRight()) ? indentation.Right.IntValue()
            //        : -1;
            throw new NotImplementedException();
        }

        /**
         * Specifies the indentation which shall be Removed from the first line of
         * the parent paragraph, by moving the indentation on the first line back
         * towards the beginning of the direction of text flow.
         * This indentation is specified relative to the paragraph indentation which is specified for
         * all other lines in the parent paragraph.
         * <p/>
         * The firstLine and hanging attributes are mutually exclusive, if both are specified, then the
         * firstLine value is ignored.
         * </p>
         *
         * @param indentation
         */

        public void SetIndentationHanging(int indentation)
        {
            //CTInd indent = GetCTInd(true);
            //Bigint bi = new Bigint("" + indentation);
            //indent.Hanging = (bi);
            throw new NotImplementedException();
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
        public int GetIndentationHanging()
        {
            //CTInd indentation = GetCTInd(false);
            //return (indentation != null && indentation.IsSetHanging()) ? indentation.Hanging.IntValue() : -1;
            throw new NotImplementedException();
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
         * value is ignored. If this attribute is omitted, then its value shall be
         * assumed to be zero (if needed).
         *
         * @param indentation
         */
        public void SetIndentationFirstLine(int indentation)
        {
            //CTInd indent = GetCTInd(true);
            //Bigint bi = new Bigint("" + indentation);
            //indent.FirstLine = (bi);
            throw new NotImplementedException();
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
        public int GetIndentationFirstLine()
        {
            //CTInd indentation = GetCTInd(false);
            //return (indentation != null && indentation.IsSetFirstLine()) ? indentation.FirstLine.IntValue()
            //        : -1;
            throw new NotImplementedException();
        }

        /**
         * This element specifies whether a consumer shall break Latin text which
         * exceeds the text extents of a line by breaking the word across two lines
         * (breaking on the character level) or by moving the word to the following
         * line (breaking on the word level).
         *
         * @param wrap - bool
         */
        public void SetWordWrap(bool wrap)
        {
            //CTOnOff wordWrap = GetCTPPr().IsSetWordWrap() ? GetCTPPr()
            //        .WordWrap : GetCTPPr().AddNewWordWrap();
            //if (wrap)
            //    wordWrap.Val = (STOnOff.TRUE);
            //else
            //    wordWrap.UnsetVal();
            throw new NotImplementedException();
        }

        /**
         * This element specifies whether a consumer shall break Latin text which
         * exceeds the text extents of a line by breaking the word across two lines
         * (breaking on the character level) or by moving the word to the following
         * line (breaking on the word level).
         *
         * @return bool
         */
        public bool IsWordWrap()
        {
            //CTOnOff wordWrap = GetCTPPr().IsSetWordWrap() ? GetCTPPr()
            //        .WordWrap : null;
            //if (wordWrap != null)
            //{
            //    return (wordWrap.Val == STOnOff.ON
            //            || wordWrap.Val == STOnOff.TRUE || wordWrap.Val == STOnOff.X_1) ? true
            //            : false;
            //}
            //return false;
            throw new NotImplementedException();
        }

        /**
         * This method provides a style to the paragraph
         * This is useful when, e.g. an Heading style has to be assigned
         * @param newStyle
         */
        public void SetStyle(String newStyle)
        {
            //CTPPr pr = GetCTPPr();
            //CTString style = pr.PStyle != null ? pr.PStyle : pr.AddNewPStyle();
            //style.Val = (newStyle);
            throw new NotImplementedException();
        }

        /**
         * @return  the style of the paragraph
         */
        public String GetStyle()
        {
            //CTPPr pr = GetCTPPr();
            //CTString style = pr.IsSetPStyle() ? pr.PStyle : null;
            //return style != null ? style.Val : null;
            throw new NotImplementedException();
        }

        /**
         * Get a <b>copy</b> of the currently used CTPBrd, if none is used, return
         * a new instance.
         */
        private CT_PBdr GetCTPBrd(bool Create)
        {
            //CTPPr pr = GetCTPPr();
            //CTPBdr ct = pr.IsSetPBdr() ? pr.PBdr : null;
            //if (create && ct == null)
            //    ct = pr.AddNewPBdr();
            //return ct;
            throw new NotImplementedException();
        }

        /**
         * Get a <b>copy</b> of the currently used CTSpacing, if none is used,
         * return a new instance.
         */
        private CT_Spacing GetCTSpacing(bool Create)
        {
            //CTPPr pr = GetCTPPr();
            //CTSpacing ct = pr.Spacing == null ? null : pr.Spacing;
            //if (create && ct == null)
            //    ct = pr.AddNewSpacing();
            //return ct;
            throw new NotImplementedException();
        }

        /**
         * Get a <b>copy</b> of the currently used CTPInd, if none is used, return
         * a new instance.
         */
        private CT_Ind GetCTInd(bool Create)
        {
            //CTPPr pr = GetCTPPr();
            //CTInd ct = pr.Ind == null ? null : pr.Ind;
            //if (create && ct == null)
            //    ct = pr.AddNewInd();
            //return ct;
            throw new NotImplementedException();
        }

        /**
         * Get a <b>copy</b> of the currently used CTPPr, if none is used, return
         * a new instance.
         */
        private CT_PPr GetCTPPr()
        {
            //CTPPr pr = paragraph.PPr == null ? paragraph.AddNewPPr()
            //        : paragraph.PPr;
            //return pr;
            throw new NotImplementedException();
        }


        /**
         * add a new run at the end of the position of 
         * the content of parameter run
         * @param run
         */
        protected void AddRun(CT_R Run){
        //int pos;
        //pos = paragraph.RList.Size();
        //paragraph.AddNewR();
        //paragraph.RArray=(pos, Run);
            throw new NotImplementedException();
    }

        /**
         * this methods parse the paragraph and search for the string searched.
         * If it Finds the string, it will return true and the position of the String
         * will be saved in the parameter startPos.
         * @param searched
         * @param startPos
         */
        public TextSegement searchText(String searched, PositionInParagraph startPos)
        {

            //int startRun = startPos.Run,
            //    startText = startPos.Text,
            //    startChar = startPos.Char;
            //int beginRunPos = 0, candCharPos = 0;
            //bool newList = false;
            //for (int RunPos = startRun; RunPos < paragraph.RList.Size(); RunPos++)
            //{
            //    int beginTextPos = 0, beginCharPos = 0, textPos = 0, charPos = 0;
            //    CTR ctRun = paragraph.GetRArray(RunPos);
            //    XmlCursor c = ctRun.NewCursor();
            //    c.SelectPath("./*");
            //    while (c.ToNextSelection())
            //    {
            //        XmlObject o = c.Object;
            //        if (o is CTText)
            //        {
            //            if (textPos >= startText)
            //            {
            //                String candidate = ((CTText)o).StringValue;
            //                if (RunPos == startRun)
            //                    charPos = startChar;
            //                else
            //                    charPos = 0;
            //                for (; charPos < candidate.Length(); charPos++)
            //                {
            //                    if ((candidate[charPos] == searched[0]) && (candCharPos == 0))
            //                    {
            //                        beginTextPos = textPos;
            //                        beginCharPos = charPos;
            //                        beginRunPos = RunPos;
            //                        newList = true;
            //                    }
            //                    if (candidate[charPos] == searched[candCharPos])
            //                    {
            //                        if (candCharPos + 1 < searched.Length())
            //                            candCharPos++;
            //                        else if (newList)
            //                        {
            //                            TextSegement segement = new TextSegement();
            //                            segement.BeginRun = (beginRunPos);
            //                            segement.BeginText = (beginTextPos);
            //                            segement.BeginChar = (beginCharPos);
            //                            segement.EndRun = (RunPos);
            //                            segement.EndText = (textPos);
            //                            segement.EndChar = (charPos);
            //                            return segement;
            //                        }
            //                    }
            //                    else
            //                        candCharPos = 0;
            //                }
            //            }
            //            textPos++;
            //        }
            //        else if (o is CTProofErr)
            //        {
            //            c.RemoveXml();
            //        }
            //        else if (o is CTRPr) ;
            //        //do nothing
            //        else
            //            candCharPos = 0;
            //    }

            //    c.Dispose();
            //}
            //return null;
            throw new NotImplementedException();
        }

        /**
         * insert a new Run in RunArray
         * @param pos
         * @return  the inserted run
         */
        public XWPFRun insertNewRun(int pos)
        {
            //if (pos >= 0 && pos <= paragraph.SizeOfRArray())
            //{
            //    CTR ctRun = paragraph.InsertNewR(pos);
            //    XWPFRun newRun = new XWPFRun(ctRun, this);
            //    Runs.Add(pos, newRun);
            //    return newRun;
            //}
            //return null;
            throw new NotImplementedException();
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
            //if (pos >= 0 && pos < paragraph.SizeOfRArray())
            //{
            //    GetCTP().RemoveR(pos);
            //    Runs.Remove(pos);
            //    return true;
            //}
            //return false;
            throw new NotImplementedException();
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
        public BodyType GetPartType()
        {
            return part.GetPartType();
        }

        /**
         * Adds a new Run to the Paragraph
         * 
         * @param r
         */
        public void AddRun(XWPFRun r)
        {
            if (!Runs.Contains(r))
            {
                Runs.Add(r);
            }
        }

        /**
         * return the XWPFRun-Element which owns the CTR Run-Element
         * 
         * @param r
         */
        public XWPFRun GetRun(CT_R r)
        {
            //for (int i = 0; i < GetRuns().Count; i++) {
            //    if (getRuns().Get(i).CTR == r) {
            //        return GetRuns().Get(i);
            //    }
            //}
            //return null;
            throw new NotImplementedException();
        }

    }//end class
}