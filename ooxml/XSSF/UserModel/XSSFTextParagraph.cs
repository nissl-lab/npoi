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
using System.Collections.Generic;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using System;
using System.Text;
using NPOI.XSSF.Model;
using System.Drawing;
namespace NPOI.XSSF.UserModel
{
    /**
     * Represents a paragraph of text within the Containing text body.
     * The paragraph is the highest level text separation mechanism.
     */
    public class XSSFTextParagraph : IEnumerator<XSSFTextRun>, IEnumerable<XSSFTextRun>
    {
        private CT_TextParagraph _p;
        private CT_Shape _shape;
        private List<XSSFTextRun> _Runs;

        public XSSFTextParagraph(CT_TextParagraph p, CT_Shape ctShape)
        {
            _p = p;
            _shape = ctShape;
            _Runs = new List<XSSFTextRun>();
            foreach (object ch in _p.r)
            {
                if (ch is CT_RegularTextRun)
                {
                    CT_RegularTextRun r = (CT_RegularTextRun)ch;
                    _Runs.Add(new XSSFTextRun(r, this));
                }
                else if (ch is CT_TextLineBreak)
                {
                    CT_TextLineBreak br = (CT_TextLineBreak)ch;
                    CT_RegularTextRun r = new CT_RegularTextRun();
                    r.rPr = (br.rPr);
                    r.t=("\n");
                    _Runs.Add(new XSSFTextRun(r, this));
                }
                else if (ch is CT_TextField)
                {
                    CT_TextField f = (CT_TextField)ch;
                    CT_RegularTextRun r = new CT_RegularTextRun();
                    r.rPr = (f.rPr);
                    r.t = (f.t);
                    _Runs.Add(new XSSFTextRun(r, this));
                }
            }
            //foreach(XmlObject ch in _p.selectPath("*")){
            //    if(ch is CTRegularTextRun){
            //        CTRegularTextRun r = (CTRegularTextRun)ch;
            //        _Runs.add(new XSSFTextRun(r, this));
            //    } else if (ch is CTTextLineBreak){
            //        CTTextLineBreak br = (CTTextLineBreak)ch;
            //        CTRegularTextRun r = CTRegularTextRun.Factory.newInstance();
            //        r.SetRPr(br.GetRPr());
            //        r.SetT("\n");
            //        _Runs.add(new XSSFTextRun(r, this));
            //    } else if (ch is CTTextField){
            //        CTTextField f = (CTTextField)ch;
            //        CTRegularTextRun r = CTRegularTextRun.Factory.newInstance();
            //        r.SetRPr(f.GetRPr());
            //        r.SetT(f.GetT());
            //        _Runs.add(new XSSFTextRun(r, this));
            //    }
            //}
        }

        public String Text
        {
            get
            {
                StringBuilder out1 = new StringBuilder();
                foreach (XSSFTextRun r in _Runs)
                {
                    out1.Append(r.Text);
                }
                return out1.ToString();
            }
        }


        public CT_TextParagraph GetXmlObject()
        {
            return _p;
        }


        public CT_Shape ParentShape
        {
            get
            {
                return _shape;
            }
        }

        public List<XSSFTextRun> TextRuns
        {
            get
            {
                return _Runs;
            }
        }

        /**
         * Add a new run of text
         *
         * @return a new run of text
         */
        public XSSFTextRun AddNewTextRun()
        {
            CT_RegularTextRun r = _p.AddNewR();
            CT_TextCharacterProperties rPr = r.AddNewRPr();
            rPr.lang = ("en-US");
            XSSFTextRun run = new XSSFTextRun(r, this);
            _Runs.Add(run);
            return run;
        }

        /**
         * Insert a line break
         *
         * @return text run representing this line break ('\n')
         */
        public XSSFTextRun AddLineBreak()
        {
            CT_TextLineBreak br = _p.AddNewBr();
            CT_TextCharacterProperties brProps = br.AddNewRPr();
            if (_Runs.Count > 0)
            {
                // by default line break has the font size of the last text run
                CT_TextCharacterProperties prevRun = _Runs[_Runs.Count - 1].GetRPr();
                brProps = (prevRun);
            }
            CT_RegularTextRun r = new CT_RegularTextRun();
            r.rPr = (brProps);
            r.t = ("\n");
            XSSFTextRun run = new XSSFLineBreak(r, this, brProps);
            _Runs.Add(run);
            return run;
        }

        /**
         * get or set the alignment that is to be applied to the paragraph.
         * Possible values for this include left, right, centered, justified and distributed,
         * If this attribute is omitted, then a value of left is implied.
         * @return alignment that is applied to the paragraph
         */
        public TextAlign TextAlign
        {
            get
            {
                ParagraphPropertyFetcher<TextAlign> fetcher = new ParagraphPropertyTextAlignFetcher(Level);
                fetchParagraphProperty(fetcher);
                //return fetcher.GetValue() == null ? TextAlign.LEFT : fetcher.GetValue();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //if(align == null) {
                //    if(pr.IsSetAlgn()) pr.unsetAlgn();
                //} else {
                //    pr.SetAlgn(STTextAlignType.Enum.forInt(align.ordinal() + 1));
                //}
                throw new NotImplementedException();
            }
        }
        private class ParagraphPropertyTextAlignFetcher : ParagraphPropertyFetcher<TextAlign>
        {
            public ParagraphPropertyTextAlignFetcher(int level) : base(level) 
            {
            }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetAlgn())
                {
                    TextAlign val = (TextAlign)(props.algn - 1); //TextAlign.values()[props.GetAlgn().intValue() - 1];
                    SetValue(val);
                    return true;
                }
                return false;
            }
        }


        /**
         * Determines where vertically on a line of text the actual words are positioned. This deals
         * with vertical placement of the characters with respect to the baselines. For instance
         * having text anchored to the top baseline, anchored to the bottom baseline, centered in
         * between, etc.
         * If this attribute is omitted, then a value of baseline is implied.
         * @return alignment that is applied to the paragraph
         */
        public TextFontAlign TextFontAlign
        {
            get
            {
                //ParagraphPropertyFetcher<TextFontAlign> fetcher = new ParagraphPropertyFetcher<TextFontAlign>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetFontAlgn()){
                //            TextFontAlign val = TextFontAlign.values()[props.GetFontAlgn().intValue() - 1];
                //            SetValue(val);
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue() == null ? TextFontAlign.BASELINE : fetcher.GetValue();        
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //if(align == null) {
                //    if(pr.IsSetFontAlgn()) pr.unsetFontAlgn();
                //} else {
                //    pr.SetFontAlgn(STTextFontAlignType.Enum.forInt(align.ordinal() + 1));
                //}
                throw new NotImplementedException();
            }
        }

        /**
         * @return the font to be used on bullet characters within a given paragraph
         */
        public String BulletFont
        {
            get
            {
                //ParagraphPropertyFetcher<String> fetcher = new ParagraphPropertyFetcher<String>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetBuFont()){
                //            SetValue(props.GetBuFont().GetTypeface());
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //CTTextFont font = pr.IsSetBuFont() ? pr.GetBuFont() : pr.AddNewBuFont();
                //font.SetTypeface(typeface);
                throw new NotImplementedException();
            }
        }

        /**
         * @return the character to be used in place of the standard bullet point
         */
        public String BulletCharacter
        {
            get
            {
                //ParagraphPropertyFetcher<String> fetcher = new ParagraphPropertyFetcher<String>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetBuChar()){
                //            SetValue(props.GetBuChar().GetChar());
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //CTTextCharBullet c = pr.IsSetBuChar() ? pr.GetBuChar() : pr.AddNewBuChar();
                //c.SetChar(str);
                throw new NotImplementedException();
            }
        }

        /**
         *
         * @return the color of bullet characters within a given paragraph.
         * A <code>null</code> value means to use the text font color.
         */
        public Color BulletFontColor
        {
            get
            {
                //ParagraphPropertyFetcher<Color> fetcher = new ParagraphPropertyFetcher<Color>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetBuClr()){
                //            if(props.GetBuClr().IsSetSrgbClr()){
                //                CTSRgbColor clr = props.GetBuClr().GetSrgbClr();
                //                byte[] rgb = clr.GetVal();
                //                SetValue(new Color(0xFF & rgb[0], 0xFF & rgb[1], 0xFF & rgb[2]));
                //                return true;
                //            }
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //CTColor c = pr.IsSetBuClr() ? pr.GetBuClr() : pr.AddNewBuClr();
                //CTSRgbColor clr = c.IsSetSrgbClr() ? c.GetSrgbClr() : c.AddNewSrgbClr();
                //clr.SetVal(new byte[]{(byte) color.GetRed(), (byte) color.GetGreen(), (byte) color.GetBlue()});
                throw new NotImplementedException();
            }
        }

        /**
         * Returns the bullet size that is to be used within a paragraph.
         * This may be specified in two different ways, percentage spacing and font point spacing:
         * <p>
         * If bulletSize >= 0, then bulletSize is a percentage of the font size.
         * If bulletSize < 0, then it specifies the size in points
         * </p>
         *
         * @return the bullet size
         */
        public double BulletFontSize
        {
            get
            {
                //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetBuSzPct()){
                //            SetValue(props.GetBuSzPct().GetVal() * 0.001);
                //            return true;
                //        }
                //        if(props.IsSetBuSzPts()){
                //            SetValue( - props.GetBuSzPts().GetVal() * 0.01);
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue() == null ? 100 : fetcher.GetValue();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();

                //if(bulletSize >= 0) {
                //    // percentage
                //    CTTextBulletSizePercent pt = pr.IsSetBuSzPct() ? pr.GetBuSzPct() : pr.AddNewBuSzPct();
                //    pt.SetVal((int)(bulletSize*1000));
                //    // unset points if percentage is now Set
                //    if(pr.IsSetBuSzPts()) pr.unsetBuSzPts();
                //} else {
                //    // points
                //    CTTextBulletSizePoint pt = pr.IsSetBuSzPts() ? pr.GetBuSzPts() : pr.AddNewBuSzPts();
                //    pt.SetVal((int)(-bulletSize*100));
                //    // unset percentage if points is now Set
                //    if(pr.IsSetBuSzPct()) pr.unsetBuSzPct();
                //}
                throw new NotImplementedException();
            }
        }

        /**
         *
         * @return the indent applied to the first line of text in the paragraph.
         */
        public double Indent
        {
            get
            {
                //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetIndent()){
                //            SetValue(Units.ToPoints(props.GetIndent()));
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);

                //return fetcher.GetValue() == null ? 0 : fetcher.GetValue();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //if(value == -1) {
                //    if(pr.IsSetIndent()) pr.unsetIndent();
                //} else {
                //    pr.SetIndent(Units.ToEMU(value));
                //}
                throw new NotImplementedException();
            }
        }



        /**
         * Specifies the left margin of the paragraph. This is specified in Addition to the text body
         * inset and applies only to this text paragraph. That is the text body inset and the LeftMargin
         * attributes are Additive with respect to the text position.
         *
         * @param value the left margin of the paragraph, -1 to clear the margin and use the default of 0.
         */
        public double LeftMargin
        {
            get
            {
                //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetMarL()){
                //            double val = Units.ToPoints(props.GetMarL());
                //            SetValue(val);
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //// if the marL attribute is omitted, then a value of 347663 is implied
                //return fetcher.GetValue() == null ? 0 : fetcher.GetValue();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //if(value == -1) {
                //    if(pr.IsSetMarL()) pr.unsetMarL();
                //} else {
                //    pr.SetMarL(Units.ToEMU(value));
                //}
                throw new NotImplementedException();
            }
        }

        /**
         * Specifies the right margin of the paragraph. This is specified in Addition to the text body
         * inset and applies only to this text paragraph. That is the text body inset and the marR
         * attributes are Additive with respect to the text position.
         *
         * @param value the right margin of the paragraph, -1 to clear the margin and use the default of 0.
         */
        public double RightMargin
        {
            get
            {
                //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetMarR()){
                //            double val = Units.ToPoints(props.GetMarR());
                //            SetValue(val);
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //// if the marL attribute is omitted, then a value of 347663 is implied
                //return fetcher.GetValue() == null ? 0 : fetcher.GetValue();        
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //if(value == -1) {
                //    if(pr.IsSetMarR()) pr.unsetMarR();
                //} else {
                //    pr.SetMarR(Units.ToEMU(value));
                //}
                throw new NotImplementedException();
            }
        }

        /**
         *
         * @return the default size for a tab character within this paragraph in points
         */
        public double DefaultTabSize
        {
            get
            {
                //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetDefTabSz()){
                //            double val = Units.ToPoints(props.GetDefTabSz());
                //            SetValue(val);
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue() == null ? 0 : fetcher.GetValue();
                throw new NotImplementedException();
            }
        }

        public double GetTabStop(int idx)
        {
            //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
            //    public bool fetch(CTTextParagraphProperties props){
            //        if(props.IsSetTabLst()){
            //            CTTextTabStopList tabStops = props.GetTabLst();
            //            if(idx < tabStops.sizeOfTabArray() ) {
            //                CTTextTabStop ts = tabStops.GetTabArray(idx);
            //                double val = Units.ToPoints(ts.GetPos());
            //                SetValue(val);
            //                return true;
            //            }
            //        }
            //        return false;
            //    }
            //};
            //fetchParagraphProperty(fetcher);
            //return fetcher.GetValue() == null ? 0. : fetcher.GetValue();
            throw new NotImplementedException();
        }
        /**
         * Add a single tab stop to be used on a line of text when there are one or more tab characters
         * present within the text. 
         * 
         * @param value the position of the tab stop relative to the left margin
         */
        public void AddTabStop(double value)
        {
            //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
            //CTTextTabStopList tabStops = pr.IsSetTabLst() ? pr.GetTabLst() : pr.AddNewTabLst();
            //tabStops.AddNewTab().SetPos(Units.ToEMU(value));
            throw new NotImplementedException();
        }

        /**
         * This element specifies the vertical line spacing that is to be used within a paragraph.
         * This may be specified in two different ways, percentage spacing and font point spacing:
         * <p>
         * If linespacing >= 0, then linespacing is a percentage of normal line height
         * If linespacing < 0, the absolute value of linespacing is the spacing in points
         * </p>
         * Examples:
         * <pre><code>
         *      // spacing will be 120% of the size of the largest text on each line
         *      paragraph.SetLineSpacing(120);
         *
         *      // spacing will be 200% of the size of the largest text on each line
         *      paragraph.SetLineSpacing(200);
         *
         *      // spacing will be 48 points
         *      paragraph.SetLineSpacing(-48.0);
         * </code></pre>
         * 
         * @param linespacing the vertical line spacing

         * Returns the vertical line spacing that is to be used within a paragraph.
         * This may be specified in two different ways, percentage spacing and font point spacing:
         * <p>
         * If linespacing >= 0, then linespacing is a percentage of normal line height.
         * If linespacing < 0, the absolute value of linespacing is the spacing in points
         * </p>
         *
         * @return the vertical line spacing.
         */
        public double LineSpacing
        {
            get
            {
                //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetLnSpc()){
                //            CTTextSpacing spc = props.GetLnSpc();

                //            if(spc.IsSetSpcPct()) SetValue( spc.GetSpcPct().GetVal()*0.001 );
                //            else if (spc.IsSetSpcPts()) SetValue( -spc.GetSpcPts().GetVal()*0.01 );
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);

                //double lnSpc = fetcher.GetValue() == null ? 100 : fetcher.GetValue();
                //if(lnSpc > 0) {
                //    // check if the percentage value is scaled
                //    CTTextNormalAutofit normAutofit = _shape.GetTxBody().GetBodyPr().GetNormAutofit();
                //    if(normAutofit != null) {
                //        double scale = 1 - (double)normAutofit.GetLnSpcReduction() / 100000;
                //        lnSpc *= scale;
                //    }
                //}

                //return lnSpc;
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //CTTextSpacing spc = CTTextSpacing.Factory.newInstance();
                //if(linespacing >= 0) spc.AddNewSpcPct().SetVal((int)(linespacing*1000));
                //else spc.AddNewSpcPts().SetVal((int)(-linespacing*100));
                //pr.SetLnSpc(spc);
                throw new NotImplementedException();
            }
        }

        /**
         * Set the amount of vertical white space that will be present before the paragraph.
         * This space is specified in either percentage or points:
         * <p>
         * If spaceBefore >= 0, then space is a percentage of normal line height.
         * If spaceBefore < 0, the absolute value of linespacing is the spacing in points
         * </p>
         * Examples:
         * <pre><code>
         *      // The paragraph will be formatted to have a spacing before the paragraph text.
         *      // The spacing will be 200% of the size of the largest text on each line
         *      paragraph.SetSpaceBefore(200);
         *
         *      // The spacing will be a size of 48 points
         *      paragraph.SetSpaceBefore(-48.0);
         * </code></pre>
         *
         * @param spaceBefore the vertical white space before the paragraph.

         * The amount of vertical white space before the paragraph
         * This may be specified in two different ways, percentage spacing and font point spacing:
         * <p>
         * If spaceBefore >= 0, then space is a percentage of normal line height.
         * If spaceBefore < 0, the absolute value of linespacing is the spacing in points
         * </p>
         *
         * @return the vertical white space before the paragraph
         */
        public double SpaceBefore
        {
            get
            {
                //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetSpcBef()){
                //            CTTextSpacing spc = props.GetSpcBef();

                //            if(spc.IsSetSpcPct()) SetValue( spc.GetSpcPct().GetVal()*0.001 );
                //            else if (spc.IsSetSpcPts()) SetValue( -spc.GetSpcPts().GetVal()*0.01 );
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);

                //double spcBef = fetcher.GetValue() == null ? 0 : fetcher.GetValue();
                //return spcBef;
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //CTTextSpacing spc = CTTextSpacing.Factory.newInstance();
                //if(spaceBefore >= 0) spc.AddNewSpcPct().SetVal((int)(spaceBefore*1000));
                //else spc.AddNewSpcPts().SetVal((int)(-spaceBefore*100));
                //pr.SetSpcBef(spc);
                throw new NotImplementedException();
            }
        }

        /**
         * Set the amount of vertical white space that will be present After the paragraph.
         * This space is specified in either percentage or points:
         * <p>
         * If spaceAfter >= 0, then space is a percentage of normal line height.
         * If spaceAfter < 0, the absolute value of linespacing is the spacing in points
         * </p>
         * Examples:
         * <pre><code>
         *      // The paragraph will be formatted to have a spacing After the paragraph text.
         *      // The spacing will be 200% of the size of the largest text on each line
         *      paragraph.SetSpaceAfter(200);
         *
         *      // The spacing will be a size of 48 points
         *      paragraph.SetSpaceAfter(-48.0);
         * </code></pre>
         *
         * @param spaceAfter the vertical white space After the paragraph.

         * The amount of vertical white space After the paragraph
         * This may be specified in two different ways, percentage spacing and font point spacing:
         * <p>
         * If spaceBefore >= 0, then space is a percentage of normal line height.
         * If spaceBefore < 0, the absolute value of linespacing is the spacing in points
         * </p>
         *
         * @return the vertical white space After the paragraph
         */
        public double SpaceAfter
        {
            get
            {
                //ParagraphPropertyFetcher<Double> fetcher = new ParagraphPropertyFetcher<Double>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetSpcAft()){
                //            CTTextSpacing spc = props.GetSpcAft();

                //            if(spc.IsSetSpcPct()) SetValue( spc.GetSpcPct().GetVal()*0.001 );
                //            else if (spc.IsSetSpcPts()) SetValue( -spc.GetSpcPts().GetVal()*0.01 );
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue() == null ? 0 : fetcher.GetValue();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
                //CTTextSpacing spc = CTTextSpacing.Factory.newInstance();
                //if(spaceAfter >= 0) spc.AddNewSpcPct().SetVal((int)(spaceAfter*1000));
                //else spc.AddNewSpcPts().SetVal((int)(-spaceAfter*100));
                //pr.SetSpcAft(spc);
                throw new NotImplementedException();
            }
        }

        /**
         * Specifies the particular level text properties that this paragraph will follow.
         * The value for this attribute formats the text according to the corresponding level
         * paragraph properties defined in the list of styles associated with the body of text
         * that this paragraph belongs to (therefore in the parent shape).
         * <p>
         * Note that the closest properties object to the text is used, therefore if there is
         * a conflict between the text paragraph properties and the list style properties for 
         * this level then the text paragraph properties will take precedence.
         * </p>
         * Returns the level of text properties that this paragraph will follow.
         * 
         * @return the text level of this paragraph (0-based). Default is 0.
         */
        public int Level
        {
            get
            {
                //CTTextParagraphProperties pr = _p.GetPPr();
                //if(pr == null) return 0;

                //return pr.GetLvl();
                throw new NotImplementedException();
            }
            set
            {
                //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();

                //pr.SetLvl(level);
                throw new NotImplementedException();
            }
        }


        /**
         * Returns whether this paragraph has bullets
         */
        public bool IsBullet()
        {
            //ParagraphPropertyFetcher<Boolean> fetcher = new ParagraphPropertyFetcher<Boolean>(getLevel()){
            //    public bool fetch(CTTextParagraphProperties props){
            //        if (props.IsSetBuNone()) {
            //            SetValue(false);
            //            return true;
            //        }
            //        if (props.IsSetBuFont()) {
            //            if (props.IsSetBuChar() || props.IsSetBuAutoNum()) {
            //                SetValue(true);
            //                return true;
            //            } else {
            //                // Excel treats text with buFont but no char/autonum
            //                //  as not bulleted
            //                // Possibly the font is just used if bullets turned on again?
            //            }
            //        }
            //        return false;
            //    }
            //};
            //fetchParagraphProperty(fetcher);
            //return fetcher.GetValue() == null ? false : fetcher.GetValue();
            throw new NotImplementedException();
        }

        /**
         * Set or unset this paragraph as a bullet point
         * 
         * @param flag whether text in this paragraph has bullets
         */
        public void SetBullet(bool flag)
        {
            //if(isBullet() == flag) return;

            //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
            //if(!flag) {
            //    pr.AddNewBuNone();

            //    if(pr.IsSetBuAutoNum()) pr.unsetBuAutoNum();
            //    if(pr.IsSetBuBlip()) pr.unsetBuBlip();
            //    if(pr.IsSetBuChar()) pr.unsetBuChar();
            //    if(pr.IsSetBuClr()) pr.unsetBuClr();
            //    if(pr.IsSetBuClrTx()) pr.unsetBuClrTx();
            //    if(pr.IsSetBuFont()) pr.unsetBuFont();
            //    if(pr.IsSetBuFontTx()) pr.unsetBuFontTx();
            //    if(pr.IsSetBuSzPct()) pr.unsetBuSzPct();
            //    if(pr.IsSetBuSzPts()) pr.unsetBuSzPts();
            //    if(pr.IsSetBuSzTx()) pr.unsetBuSzTx();
            //} else {
            //    if(pr.IsSetBuNone()) pr.unsetBuNone();
            //    if(!pr.IsSetBuFont()) pr.AddNewBuFont().SetTypeface("Arial");
            //    if(!pr.IsSetBuAutoNum()) pr.AddNewBuChar().SetChar("\u2022");
            //}
            throw new NotImplementedException();
        }

        /**
         * Set this paragraph as an automatic numbered bullet point
         *
         * @param scheme type of auto-numbering
         * @param startAt the number that will start number for a given sequence of automatically
         *        numbered bullets (1-based).
         */
        public void SetBullet(ListAutoNumber scheme, int startAt)
        {
            //if(startAt < 1) throw new ArgumentException("Start Number must be greater or equal that 1") ;
            //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
            //CTTextAutonumberBullet lst = pr.IsSetBuAutoNum() ? pr.GetBuAutoNum() : pr.AddNewBuAutoNum();        
            //lst.SetType(STTextAutonumberScheme.Enum.forInt(scheme.ordinal() + 1));
            //lst.SetStartAt(startAt);

            //if(!pr.IsSetBuFont()) pr.AddNewBuFont().SetTypeface("Arial");
            //if(pr.IsSetBuNone()) pr.unsetBuNone();        
            //// remove these elements if present as it results in invalid content when opening in Excel.
            //if(pr.IsSetBuBlip()) pr.unsetBuBlip();
            //if(pr.IsSetBuChar()) pr.unsetBuChar();        
            throw new NotImplementedException();
        }

        /**
         * Set this paragraph as an automatic numbered bullet point
         *
         * @param scheme type of auto-numbering
         */
        public void SetBullet(ListAutoNumber scheme)
        {
            //CTTextParagraphProperties pr = _p.IsSetPPr() ? _p.GetPPr() : _p.AddNewPPr();
            //CTTextAutonumberBullet lst = pr.IsSetBuAutoNum() ? pr.GetBuAutoNum() : pr.AddNewBuAutoNum();
            //lst.SetType(STTextAutonumberScheme.Enum.forInt(scheme.ordinal() + 1));

            //if(!pr.IsSetBuFont()) pr.AddNewBuFont().SetTypeface("Arial");
            //if(pr.IsSetBuNone()) pr.unsetBuNone();
            //// remove these elements if present as it results in invalid content when opening in Excel.
            //if(pr.IsSetBuBlip()) pr.unsetBuBlip();
            //if(pr.IsSetBuChar()) pr.unsetBuChar();
            throw new NotImplementedException();
        }

        /**
         * Returns whether this paragraph has automatic numbered bullets
         */
        public bool IsBulletAutoNumber
        {
            get
            {
                //ParagraphPropertyFetcher<Boolean> fetcher = new ParagraphPropertyFetcher<Boolean>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetBuAutoNum()) {
                //            SetValue(true);
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue() == null ? false : fetcher.GetValue();
                throw new NotImplementedException();
            }
        }

        /**
         * Returns the starting number if this paragraph has automatic numbered bullets, otherwise returns 0
         */
        public int BulletAutoNumberStart
        {
            get
            {
                //ParagraphPropertyFetcher<int> fetcher = new ParagraphPropertyFetcher<int>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetBuAutoNum() && props.GetBuAutoNum().IsSetStartAt()) {
                //            SetValue(props.GetBuAutoNum().GetStartAt());
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);
                //return fetcher.GetValue() == null ? 0 : fetcher.GetValue();
                throw new NotImplementedException();
            }
        }

        /**
         * Returns the auto number scheme if this paragraph has automatic numbered bullets, otherwise returns ListAutoNumber.ARABIC_PLAIN
         */
        public ListAutoNumber BulletAutoNumberScheme
        {
            get
            {
                //ParagraphPropertyFetcher<ListAutoNumber> fetcher = new ParagraphPropertyFetcher<ListAutoNumber>(getLevel()){
                //    public bool fetch(CTTextParagraphProperties props){
                //        if(props.IsSetBuAutoNum()) {
                //            SetValue(ListAutoNumber.values()[props.GetBuAutoNum().GetType().intValue() - 1]);
                //            return true;
                //        }
                //        return false;
                //    }
                //};
                //fetchParagraphProperty(fetcher);

                //// Note: documentation does not define a default, return ListAutoNumber.ARABIC_PLAIN (1,2,3...)
                //return fetcher.GetValue() == null ? ListAutoNumber.ARABIC_PLAIN : fetcher.GetValue();
                throw new NotImplementedException();
            }
        }


        private bool fetchParagraphProperty(ParagraphPropertyFetcher visitor)
        {
            bool ok = false;

            if (_p.IsSetPPr()) ok = visitor.Fetch(_p.pPr);

            if (!ok)
            {
                ok = visitor.Fetch(_shape);
            }

            return ok;
        }


        public override String ToString()
        {
            return "[" + this.GetType().ToString() + "]" + Text;
        }

        public XSSFTextRun Current
        {
            get { throw new NotImplementedException(); }
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        object System.Collections.IEnumerator.Current
        {
            get { throw new NotImplementedException(); }
        }

        public bool MoveNext()
        {
            throw new NotImplementedException();
        }

        public void Reset()
        {
            throw new NotImplementedException();
        }

        public IEnumerator<XSSFTextRun> GetEnumerator()
        {
            return _Runs.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _Runs.GetEnumerator();
        }
    }

}