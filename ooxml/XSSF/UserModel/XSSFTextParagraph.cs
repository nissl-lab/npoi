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
using NPOI.Util;
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
                ParagraphPropertyTextAlignFetcher fetcher = new ParagraphPropertyTextAlignFetcher(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue() == null ? TextAlign.LEFT : fetcher.GetValue().Value;
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                if (value == TextAlign.None)
                {
                    if (pr.IsSetAlgn()) pr.UnsetAlgn();
                }
                else
                {
                    pr.algn = (ST_TextAlignType)(value - 1);
                }
            }
        }
        private class ParagraphPropertyTextAlignFetcher : ParagraphPropertyFetcher<TextAlign?>
        {
            public ParagraphPropertyTextAlignFetcher(int level) : base(level) 
            {
            }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetAlgn())
                {
                    TextAlign val = (TextAlign)(props.algn + 1); //TextAlign.values()[props.GetAlgn().intValue() - 1];
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
                ParagraphPropertyFetcherTextFontAlign fetcher = new ParagraphPropertyFetcherTextFontAlign(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue() == null ? TextFontAlign.BASELINE : fetcher.GetValue().Value;        
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                if (value == TextFontAlign.None)
                {
                    if (pr.IsSetFontAlgn()) pr.UnsetFontAlgn();
                }
                else
                {
                    pr.fontAlgn = (ST_TextFontAlignType)(value - 1);
                }
            }
        }
        class ParagraphPropertyFetcherTextFontAlign : ParagraphPropertyFetcher<TextFontAlign?>
        {
            public ParagraphPropertyFetcherTextFontAlign(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetFontAlgn())
                {
                    TextFontAlign val = (TextFontAlign)(props.fontAlgn + 1);
                    SetValue(val);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherBulletFont fetcher = new ParagraphPropertyFetcherBulletFont(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue();
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                CT_TextFont font = pr.IsSetBuFont() ? pr.buFont : pr.AddNewBuFont();
                font.typeface = (value);
            }
        }
        class ParagraphPropertyFetcherBulletFont : ParagraphPropertyFetcher<String>
        {
            public ParagraphPropertyFetcherBulletFont(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetBuFont())
                {
                    SetValue(props.buFont.typeface);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherBulletCharacter fetcher = new ParagraphPropertyFetcherBulletCharacter(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue();
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                CT_TextCharBullet c = pr.IsSetBuChar() ? pr.buChar : pr.AddNewBuChar();
                c.@char = (value);
            }
        }
        class ParagraphPropertyFetcherBulletCharacter : ParagraphPropertyFetcher<String>
        {
            public ParagraphPropertyFetcherBulletCharacter(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetBuChar())
                {
                    SetValue(props.buChar.@char);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherBulletFontColor fetcher = new ParagraphPropertyFetcherBulletFontColor(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue();
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                CT_Color c = pr.IsSetBuClr() ? pr.buClr : pr.AddNewBuClr();
                CT_SRgbColor clr = c.IsSetSrgbClr() ? c.srgbClr : c.AddNewSrgbClr();
                clr.val = (new byte[] { value.R, value.G, value.B });
            }
        }
        class ParagraphPropertyFetcherBulletFontColor : ParagraphPropertyFetcher<Color>
        {
            public ParagraphPropertyFetcherBulletFontColor(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetBuClr())
                {
                    if (props.buClr.IsSetSrgbClr())
                    {
                        CT_SRgbColor clr = props.buClr.srgbClr;
                        byte[] rgb = clr.val;
                        SetValue(Color.FromArgb(0xFF & rgb[0], 0xFF & rgb[1], 0xFF & rgb[2]));
                        return true;
                    }
                }
                return false;
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
                ParagraphPropertyFetcherBulletFontSize fetcher = new ParagraphPropertyFetcherBulletFontSize(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue() == null ? 100 : fetcher.GetValue().Value;
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();

                if (value >= 0)
                {
                    // percentage
                    CT_TextBulletSizePercent pt = pr.IsSetBuSzPct() ? pr.buSzPct : pr.AddNewBuSzPct();
                    pt.val = ((int)(value * 1000));
                    // unset points if percentage is now Set
                    if (pr.IsSetBuSzPts()) pr.UnsetBuSzPts();
                }
                else
                {
                    // points
                    CT_TextBulletSizePoint pt = pr.IsSetBuSzPts() ? pr.buSzPts : pr.AddNewBuSzPts();
                    pt.val = ((int)(-value * 100));
                    // unset percentage if points is now Set
                    if (pr.IsSetBuSzPct()) pr.UnsetBuSzPct();
                }
            }
        }
        class ParagraphPropertyFetcherBulletFontSize : ParagraphPropertyFetcher<double?>
        {
            public ParagraphPropertyFetcherBulletFontSize(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetBuSzPct())
                {
                    SetValue(props.buSzPct.val * 0.001);
                    return true;
                }
                if (props.IsSetBuSzPts())
                {
                    SetValue(-props.buSzPts.val * 0.01);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherIndent fetcher = new ParagraphPropertyFetcherIndent(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue();
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                if (value == -1)
                {
                    if (pr.IsSetIndent()) pr.UnsetIndent();
                }
                else
                {
                    pr.indent = (Units.ToEMU(value));
                }
            }
        }

        class ParagraphPropertyFetcherIndent : ParagraphPropertyFetcher<double>
        {
            public ParagraphPropertyFetcherIndent(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetIndent())
                {
                    SetValue(Units.ToPoints(props.indent));
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherLeftMargin fetcher = new ParagraphPropertyFetcherLeftMargin(Level);
                fetchParagraphProperty(fetcher);
                // if the marL attribute is omitted, then a value of 347663 is implied
                return fetcher.GetValue();
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                if (value == -1)
                {
                    if (pr.IsSetMarL()) pr.UnsetMarL();
                }
                else
                {
                    pr.marL = (Units.ToEMU(value));
                }
            }
        }
        class ParagraphPropertyFetcherLeftMargin : ParagraphPropertyFetcher<double>
        {
            public ParagraphPropertyFetcherLeftMargin(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetMarL())
                {
                    double val = Units.ToPoints(props.marL);
                    SetValue(val);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherRightMargin fetcher = new ParagraphPropertyFetcherRightMargin(Level);
                fetchParagraphProperty(fetcher);
                // if the marL attribute is omitted, then a value of 347663 is implied
                return fetcher.GetValue();        
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                if (value == -1)
                {
                    if (pr.IsSetMarR()) pr.UnsetMarR();
                }
                else
                {
                    pr.marR = (Units.ToEMU(value));
                }
            }
        }
        class ParagraphPropertyFetcherRightMargin : ParagraphPropertyFetcher<double>
        {
            public ParagraphPropertyFetcherRightMargin(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetMarR())
                {
                    double val = Units.ToPoints(props.marR);
                    SetValue(val);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherDefaultTabSize fetcher = new ParagraphPropertyFetcherDefaultTabSize(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue();
            }
        }
        class ParagraphPropertyFetcherDefaultTabSize : ParagraphPropertyFetcher<double>
        {
            public ParagraphPropertyFetcherDefaultTabSize(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetDefTabSz())
                {
                    double val = Units.ToPoints(props.defTabSz);
                    SetValue(val);
                    return true;
                }
                return false;
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
            ParagraphPropertyFetcherTabStop fetcher = new ParagraphPropertyFetcherTabStop(Level, idx);
            fetchParagraphProperty(fetcher);
            return fetcher.GetValue();
        }

        class ParagraphPropertyFetcherTabStop : ParagraphPropertyFetcher<double>
        {
            private int idx;
            public ParagraphPropertyFetcherTabStop(int level, int idx) : base(level) { this.idx = idx; }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetTabLst())
                {
                    CT_TextTabStopList tabStops = props.tabLst;
                    if (idx < tabStops.SizeOfTabArray())
                    {
                        CT_TextTabStop ts = tabStops.GetTabArray(idx);
                        double val = Units.ToPoints(ts.pos);
                        SetValue(val);
                        return true;
                    }
                }
                return false;
            }
        }
        /**
         * Add a single tab stop to be used on a line of text when there are one or more tab characters
         * present within the text. 
         * 
         * @param value the position of the tab stop relative to the left margin
         */
        public void AddTabStop(double value)
        {
            CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
            CT_TextTabStopList tabStops = pr.IsSetTabLst() ? pr.tabLst : pr.AddNewTabLst();
            tabStops.AddNewTab().pos = (Units.ToEMU(value));
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
                ParagraphPropertyFetcherLineSpacing fetcher = new ParagraphPropertyFetcherLineSpacing(Level);
                fetchParagraphProperty(fetcher);

                double lnSpc = fetcher.GetValue() == null ? 100 : fetcher.GetValue().Value;
                if (lnSpc > 0)
                {
                    // check if the percentage value is scaled
                    CT_TextNormalAutofit normAutofit = _shape.txBody.bodyPr.normAutofit;
                    if (normAutofit != null)
                    {
                        double scale = 1 - (double)normAutofit.lnSpcReduction / 100000;
                        lnSpc *= scale;
                    }
                }

                return lnSpc;
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                CT_TextSpacing spc = new CT_TextSpacing();
                if (value >= 0) spc.AddNewSpcPct().val = ((int)(value * 1000));
                else spc.AddNewSpcPts().val = ((int)(-value * 100));
                pr.lnSpc = (spc);
            }
        }
        class ParagraphPropertyFetcherLineSpacing : ParagraphPropertyFetcher<double?>
        {
            public ParagraphPropertyFetcherLineSpacing(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetLnSpc())
                {
                    CT_TextSpacing spc = props.lnSpc;

                    if (spc.IsSetSpcPct()) SetValue(spc.spcPct.val * 0.001);
                    else if (spc.IsSetSpcPts()) SetValue(-spc.spcPts.val * 0.01);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherSpaceBefore fetcher = new ParagraphPropertyFetcherSpaceBefore(Level);
                fetchParagraphProperty(fetcher);

                double spcBef = fetcher.GetValue();
                return spcBef;
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                CT_TextSpacing spc = new CT_TextSpacing();
                if (value >= 0) spc.AddNewSpcPct().val = ((int)(value * 1000));
                else spc.AddNewSpcPts().val = ((int)(-value * 100));
                pr.spcBef = (spc);
            }
        }
        class ParagraphPropertyFetcherSpaceBefore : ParagraphPropertyFetcher<double>
        {
            public ParagraphPropertyFetcherSpaceBefore(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if(props.IsSetSpcBef()){
                    CT_TextSpacing spc = props.spcBef;

                    if (spc.IsSetSpcPct()) SetValue(spc.spcPct.val * 0.001);
                    else if (spc.IsSetSpcPts()) SetValue(-spc.spcPts.val * 0.01);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherSpaceAfter fetcher = new ParagraphPropertyFetcherSpaceAfter(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue();
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                CT_TextSpacing spc = new CT_TextSpacing();
                if (value >= 0) spc.AddNewSpcPct().val = ((int)(value * 1000));
                else spc.AddNewSpcPts().val = ((int)(-value * 100));
                pr.spcAft = (spc);
            }
        }
        class ParagraphPropertyFetcherSpaceAfter : ParagraphPropertyFetcher<double>
        {
            public ParagraphPropertyFetcherSpaceAfter(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetSpcAft())
                {
                    CT_TextSpacing spc = props.spcAft;

                    if (spc.IsSetSpcPct()) SetValue(spc.spcPct.val * 0.001);
                    else if (spc.IsSetSpcPts()) SetValue(-spc.spcPts.val * 0.01);
                    return true;
                }
                return false;
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
                CT_TextParagraphProperties pr = _p.pPr;
                if(pr == null) return 0;

                return pr.lvl;
            }
            set
            {
                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();

                pr.lvl = (value);
            }
        }


        /**
         * Returns whether this paragraph has bullets
         */
        public bool IsBullet
        {
            get
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
                ParagraphPropertyFetcherBullet fetcher = new ParagraphPropertyFetcherBullet(this.Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue() == null ? false : fetcher.GetValue().Value;
            }
            set
            {
                if(IsBullet == value) return;

                CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
                if (!value)
                {
                    pr.AddNewBuNone();

                    if (pr.IsSetBuAutoNum()) pr.UnsetBuAutoNum();
                    if (pr.IsSetBuBlip()) pr.UnsetBuBlip();
                    if (pr.IsSetBuChar()) pr.UnsetBuChar();
                    if (pr.IsSetBuClr()) pr.UnsetBuClr();
                    if (pr.IsSetBuClrTx()) pr.UnsetBuClrTx();
                    if (pr.IsSetBuFont()) pr.UnsetBuFont();
                    if (pr.IsSetBuFontTx()) pr.UnsetBuFontTx();
                    if (pr.IsSetBuSzPct()) pr.UnsetBuSzPct();
                    if (pr.IsSetBuSzPts()) pr.UnsetBuSzPts();
                    if (pr.IsSetBuSzTx()) pr.UnsetBuSzTx();
                }
                else
                {
                    if (pr.IsSetBuNone()) pr.UnsetBuNone();
                    if (!pr.IsSetBuFont()) pr.AddNewBuFont().typeface = ("Arial");
                    if (!pr.IsSetBuAutoNum()) pr.AddNewBuChar().@char = ("\u2022");
                }
            }
        }
        class ParagraphPropertyFetcherBullet : ParagraphPropertyFetcher<bool?>
        {
            public ParagraphPropertyFetcherBullet(int level) : base(level){}
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetBuNone())
                {
                    SetValue(false);
                    return true;
                }
                if (props.IsSetBuFont())
                {
                    if (props.IsSetBuChar() || props.IsSetBuAutoNum())
                    {
                        SetValue(true);
                        return true;
                    }
                    else
                    {
                        // Excel treats text with buFont but no char/autonum
                        //  as not bulleted
                        // Possibly the font is just used if bullets turned on again?
                    }
                }
                return false;
            }
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
            if (startAt < 1) throw new ArgumentException("Start Number must be greater or equal that 1");
            CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
            CT_TextAutonumberBullet lst = pr.IsSetBuAutoNum() ? pr.buAutoNum : pr.AddNewBuAutoNum();
            lst.type = (ST_TextAutonumberScheme)(int)scheme;
            lst.startAt = (startAt);

            if (!pr.IsSetBuFont()) pr.AddNewBuFont().typeface = ("Arial");
            if (pr.IsSetBuNone()) pr.UnsetBuNone();
            // remove these elements if present as it results in invalid content when opening in Excel.
            if (pr.IsSetBuBlip()) pr.UnsetBuBlip();
            if (pr.IsSetBuChar()) pr.UnsetBuChar();        
        }

        /**
         * Set this paragraph as an automatic numbered bullet point
         *
         * @param scheme type of auto-numbering
         */
        public void SetBullet(ListAutoNumber scheme)
        {
            CT_TextParagraphProperties pr = _p.IsSetPPr() ? _p.pPr : _p.AddNewPPr();
            CT_TextAutonumberBullet lst = pr.IsSetBuAutoNum() ? pr.buAutoNum : pr.AddNewBuAutoNum();
            lst.type = (ST_TextAutonumberScheme)(int)scheme;

            if (!pr.IsSetBuFont()) pr.AddNewBuFont().typeface = ("Arial");
            if (pr.IsSetBuNone()) pr.UnsetBuNone();
            // remove these elements if present as it results in invalid content when opening in Excel.
            if (pr.IsSetBuBlip()) pr.UnsetBuBlip();
            if (pr.IsSetBuChar()) pr.UnsetBuChar();
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
                ParagraphPropertyFetcherIsBulletAutoNumber fetcher = new ParagraphPropertyFetcherIsBulletAutoNumber(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue();
            }
        }
        class ParagraphPropertyFetcherIsBulletAutoNumber : ParagraphPropertyFetcher<bool>
        {
            public ParagraphPropertyFetcherIsBulletAutoNumber(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetBuAutoNum())
                {
                    SetValue(true);
                    return true;
                }
                return false;
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
                ParagraphPropertyFetcherBulletAutoNumberStart fetcher =
                    new ParagraphPropertyFetcherBulletAutoNumberStart(Level);
                fetchParagraphProperty(fetcher);
                return fetcher.GetValue();
            }
        }
        class ParagraphPropertyFetcherBulletAutoNumberStart : ParagraphPropertyFetcher<int>
        {
            public ParagraphPropertyFetcherBulletAutoNumberStart(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetBuAutoNum() && props.buAutoNum.IsSetStartAt())
                {
                    SetValue(props.buAutoNum.startAt);
                    return true;
                }
                return false;
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

                ParagraphPropertyFetcherBulletAutoNumberScheme fetcher =
                    new ParagraphPropertyFetcherBulletAutoNumberScheme(Level);
                fetchParagraphProperty(fetcher);

                // Note: documentation does not define a default, return ListAutoNumber.ARABIC_PLAIN (1,2,3...)
                return fetcher.GetValue() == null ? ListAutoNumber.ARABIC_PLAIN : fetcher.GetValue().Value;
            }
        }

        class ParagraphPropertyFetcherBulletAutoNumberScheme : ParagraphPropertyFetcher<ListAutoNumber?>
        {
            public ParagraphPropertyFetcherBulletAutoNumberScheme(int level) : base(level) { }
            public override bool Fetch(CT_TextParagraphProperties props)
            {
                if (props.IsSetBuAutoNum())
                {
                    SetValue((ListAutoNumber)(int)props.buAutoNum.type);
                    return true;
                }
                return false;
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