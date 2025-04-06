/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel.Text
{
    using NPOI.Util;
    using NPOI.XDDF.UserModel;
    using NPOI.OpenXmlFormats.Dml;
    using NPOI.Util.Optional;
    using SixLabors.Fonts;

    /// <summary>
    /// Represents a paragraph of text within the containing text body. The paragraph
    /// is the highest level text separation mechanism.
    /// </summary>
    public class XDDFTextParagraph
    {
        private XDDFTextBody _parent;
        private XDDFParagraphProperties _properties;
        private  CT_TextParagraph _p;
        private  List<XDDFTextRun> _runs;
        internal XDDFTextParagraph(CT_TextParagraph paragraph, XDDFTextBody parent)
        {
            this._p = paragraph;
            this._parent = parent;

            int count = paragraph.SizeOfBrArray() + paragraph.SizeOfFldArray() + paragraph.SizeOfRArray();
            this._runs = new List<XDDFTextRun>(count);

            foreach(object xo in _p.items)
            {
                if(xo is CT_TextLineBreak)
                {
                    _runs.Add(new XDDFTextRun((CT_TextLineBreak) xo, this));
                }
                else if(xo is CT_TextField)
                {
                    _runs.Add(new XDDFTextRun((CT_TextField) xo, this));
                }
                else if(xo is CT_RegularTextRun)
                {
                    _runs.Add(new XDDFTextRun((CT_RegularTextRun) xo, this));
                }
            }

            AddDefaultRunProperties();
            AddAfterLastRunProperties();
        }

        private void SetText(string text)
        {
            // remove all runs
            for(int i = _p.SizeOfBrArray() - 1; i >= 0; i--)
            {
                _p.RemoveBr(i);
            }
            for(int i = _p.SizeOfFldArray() - 1; i >= 0; i--)
            {
                _p.RemoveFld(i);
            }
            for(int i = _p.SizeOfRArray() - 1; i >= 0; i--)
            {
                _p.RemoveR(i);
            }
            _runs.Clear();
            AppendRegularRun(text);
        }

        private string GetText()
        {
            StringBuilder out1 = new StringBuilder();
            foreach(XDDFTextRun r in _runs)
            {
                out1.Append(r.Text);
            }
            return out1.ToString();
        }

        public string Text
        {
            get
            {
                return GetText();
            }
            set
            {
                SetText(value);
            }
        }

        public XDDFTextBody GetParentBody()
        {
            return _parent;
        }

        public List<XDDFTextRun> GetTextRuns()
        {
            return _runs;
        }

        public IEnumerator<XDDFTextRun> GetEnumerator()
        {
            return _runs.GetEnumerator();
        }

        /// <summary>
        /// Append a line break.
        /// </summary>
        /// <returns>text run representing this line break ('\n').</returns>
        public XDDFTextRun AppendLineBreak()
        {
            CT_TextLineBreak br = _p.AddNewBr();
            // by default, line break has the font properties of the last text run
            _runs.Reverse();
            foreach(XDDFTextRun tr in _runs)
            {
                CT_TextCharacterProperties prevProps = tr.GetProperties();
                // let's find one that is not undefined
                if(prevProps != null)
                {
                    br.rPr = prevProps.Copy();
                    break;
                }
            }
            XDDFTextRun run = new XDDFTextRun(br, this);
            _runs.Add(run);
            return run;
        }

        /// <summary>
        /// Append a new text field.
        /// </summary>
        /// <returns>the new text field.</returns>
        public XDDFTextRun AppendField(string id, string type, string text)
        {
            CT_TextField f = _p.AddNewFld();
            f.id = id;
            f.type = type;
            f.t = text;
            CT_TextCharacterProperties rPr = f.AddNewRPr();
            rPr.lang = LocaleUtil.GetUserLocale().Name;
            XDDFTextRun run = new XDDFTextRun(f, this);
            _runs.Add(run);
            return run;
        }

        /// <summary>
        /// Append a new run of text.
        /// </summary>
        /// <returns>the new run of text.</returns>
        public XDDFTextRun AppendRegularRun(string text)
        {
            CT_RegularTextRun r = _p.AddNewR();
            r.t = text;
            CT_TextCharacterProperties rPr = r.AddNewRPr();
            rPr.lang = LocaleUtil.GetUserLocale().Name;
            XDDFTextRun run = new XDDFTextRun(r, this);
            _runs.Add(run);
            return run;
        }

        /// <summary>
        /// <para>
        /// Get or set the alignment that is applied to the paragraph.Possible
        /// values for this include left, right, centered, justified and distributed,
        /// </para>
        /// <para>
        /// If this attribute is omitted, then a value of left is implied.
        /// </para>
        /// </summary>
        public TextAlignment? TextAlignment
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.IsSetAlgn(), props => props.algn)
                    .MapValue(align => TextAlignmentExtensions.ValueOf(align))
                    .OrElse(null);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetTextAlignment(value);
                }
            }
        }

        /// <summary>
        /// Determines where vertically on a line of text the actual words are
        /// positioned. This deals with vertical placement of the characters with
        /// respect to the baselines. For instance having text anchored to the top
        /// baseline, anchored to the bottom baseline, centered in between, etc.
        /// </summary>
        public FontAlignment? FontAlignment
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.IsSetFontAlgn(), props => props.fontAlgn)
                    .MapValue(align => FontAlignmentExtensions.ValueOf(align)).OrElse(null);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetFontAlignment(value.Value);
                }
            }
        }

        /// <summary>
        /// Get or set the indentation size that will be applied to the first line of
        /// text in the paragraph.
        /// 
        /// the indentation in points. The value <c>null</c> unsets
        /// the indentation for this paragraph.
        /// <dl>
        /// <dt>Minimum inclusive =</dt>
        /// <dd>-4032</dd>
        /// <dt>Maximum inclusive =</dt>
        /// <dd>4032</dd></dt>
        /// </summary>
        public double? Indentation
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.IsSetIndent(), props => props.indent)
                    .MapValue(emu => Units.ToPoints(emu)).OrElse(null);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetIndentation(value.Value);
                }
            }
        }

        /// <summary>
        /// Get or set the left margin of the paragraph. This is specified in addition
        /// to the text body inset and applies only to this text paragraph. That is
        /// the text body inset and the LeftMargin attributes are additive with
        /// respect to the text position.
        /// 
        /// the margin in points. The value <c>null</c> unsets the
        /// left margin for this paragraph.
        /// <dl>
        /// <dt>Minimum inclusive =</dt>
        /// <dd>0</dd>
        /// <dt>Maximum inclusive =</dt>
        /// <dd>4032</dd></dt>
        /// </summary>
        public double? MarginLeft
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.IsSetMarL(), props => props.marL)
                .MapValue(emu => Units.ToPoints(emu)).OrElse(null);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetMarginLeft(value.Value);
                }
            }

        }

        /// <summary>
        /// Get or set the right margin of the paragraph. This is specified in
        /// addition to the text body inset and applies only to this text paragraph.
        /// That is the text body inset and the RightMargin attributes are additive
        /// with respect to the text position.
        /// </summary>
        /// <param name="points">
        /// the margin in points. The value <c>null</c> unsets the
        /// right margin for this paragraph.
        /// <dl>
        /// <dt>Minimum inclusive =</dt>
        /// <dd>0</dd>
        /// <dt>Maximum inclusive =</dt>
        /// <dd>4032</dd></dt>
        /// </param>
        public double? MarginRight
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.IsSetMarR(), props => props.marR)
                .MapValue(emu => Units.ToPoints(emu)).OrElse(null);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetMarginRight(value.Value);
                }
            }
        }

        /// <summary>
        /// <para>
        /// Get or set the default size for a tab character within this paragraph.
        /// </para>
        /// <para>
        /// the default tab size in points. The value <c>null</c>
        /// unsets the default tab size for this paragraph.
        /// </para>
        /// </summary>
        public double? DefaultTabSize
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.IsSetDefTabSz(), props => props.defTabSz)
                    .MapValue(emu => Units.ToPoints(emu)).OrElse(null);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetDefaultTabSize(value.Value);
                }
            }
        }

        /// <summary>
        /// <para>
        /// This element specifies the vertical line spacing that is to be used
        /// within a paragraph. This may be specified in two different ways,
        /// percentage spacing or font points spacing:
        /// </para>
        /// <para>
        /// If spacing is instance of XDDFSpacingPercent, then line spacing is a
        /// percentage of normal line height. If spacing is instance of
        /// XDDFSpacingPoints, then line spacing is expressed in points.
        /// </para>
        /// <para>
        /// Examples:
        /// </para>
        /// <para>
        /// <code>
        ///      // spacing will be 120% of the size of the largest text on each line
        ///      paragraph.SetLineSpacing(new XDDFSpacingPercent(120));
        ///      // spacing will be 200% of the size of the largest text on each line
        ///      paragraph.SetLineSpacing(new XDDFSpacingPercent(200));
        ///      // spacing will be 48 points
        ///      paragraph.SetLineSpacing(new XDDFSpacingPoints(48.0));
        /// </code>
        /// </para>
        /// </summary>

        public XDDFSpacing LineSpacing
        {
            get
            {
                return FindDefinedParagraphProperty(props => props.IsSetLnSpc(), props => props.lnSpc)
                    .Map(spacing => ExtractSpacing(spacing)).OrElse(null);
            }
            set
            {
                if(value != null || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetLineSpacing(value);
                }
            }
        }

        /// <summary>
        /// <para>
        /// Set the amount of vertical white space that will be present before the
        /// paragraph. This may be specified in two different ways, percentage
        /// spacing or font points spacing:
        /// </para>
        /// <para>
        /// If spacing is instance of XDDFSpacingPercent, then spacing is a
        /// percentage of normal line height. If spacing is instance of
        /// XDDFSpacingPoints, then spacing is expressed in points.
        /// </para>
        /// <para>
        /// Examples:
        /// </para>
        /// <para>
        /// <code>
        ///      // The paragraph will be formatted to have a spacing before the paragraph text.
        ///      // The spacing will be 200% of the size of the largest text on each line
        ///      paragraph.SetSpaceBefore(new XDDFSpacingPercent(200));
        ///      // The spacing will be a size of 48 points
        ///      paragraph.SetSpaceBefore(new XDDFSpacingPoints(48.0));
        /// </code>
        /// </para>
        /// </summary>
        public XDDFSpacing SpaceBefore
        {
            get
            {
                return FindDefinedParagraphProperty(props => props.IsSetSpcBef(), props => props.spcBef)
                    .Map(spacing => ExtractSpacing(spacing)).OrElse(null);
            }
            set
            {
                if(value != null || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetSpaceBefore(value);
                }
            }
        }

        /// <summary>
        /// <para>
        /// Set the amount of vertical white space that will be present After the
        /// paragraph. This may be specified in two different ways, percentage
        /// spacing or font points spacing:
        /// </para>
        /// <para>
        /// If spacing is instance of XDDFSpacingPercent, then spacing is a
        /// percentage of normal line height. If spacing is instance of
        /// XDDFSpacingPoints, then spacing is expressed in points.
        /// </para>
        /// <para>
        /// Examples:
        /// </para>
        /// <para>
        /// <code>
        ///      // The paragraph will be formatted to have a spacing After the paragraph text.
        ///      // The spacing will be 200% of the size of the largest text on each line
        ///      paragraph.SetSpaceAfter(new XDDFSpacingPercent(200));
        ///      // The spacing will be a size of 48 points
        ///      paragraph.SetSpaceAfter(new XDDFSpacingPoints(48.0));
        /// </code>
        /// </para>
        /// </summary>
        public XDDFSpacing SpaceAfter
        {
            get
            {
                return FindDefinedParagraphProperty(props => props.IsSetSpcAft(), props => props.spcAft)
                    .Map(spacing => ExtractSpacing(spacing)).OrElse(null);
            }
            set
            {
                if(value != null || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetSpaceAfter(value);
                }
            }
        }

        /// <summary>
        /// Get or set the color to be used on bullet characters within a given paragraph.
        /// </summary>
        public XDDFColor BulletColor
        {
            get
            {
                return FindDefinedParagraphProperty(props => props.IsSetBuClr() || props.IsSetBuClrTx(),
                    props => new XDDFParagraphBulletProperties(props).BulletColor).OrElse(null);
            }
            set
            {
                if(value != null || _p.IsSetPPr())
                {
                    GetOrCreateBulletProperties().BulletColor = value;
                }
            }

        }

        /// <summary>
        /// Specifies the color to be used on bullet characters has to follow text
        /// color within a given paragraph.
        /// </summary>
        public void SetBulletColorFollowText()
        {
            GetOrCreateBulletProperties().SetBulletColorFollowText();
        }

        /// <summary>
        /// Get or set the font to be used on bullet characters within a given paragraph.
        /// </summary>
        public XDDFFont BulletFont
        {
            get
            {
                return FindDefinedParagraphProperty(props => props.IsSetBuFont() || props.IsSetBuFontTx(),
                    props => new XDDFParagraphBulletProperties(props).BulletFont).OrElse(null);
            }
            set
            {
                if(value != null || _p.IsSetPPr())
                {
                    GetOrCreateBulletProperties().BulletFont = value;
                }
            }
        }


        /// <summary>
        /// Specifies the font to be used on bullet characters has to follow text
        /// font within a given paragraph.
        /// </summary>
        public void SetBulletFontFollowText()
        {
            GetOrCreateBulletProperties().SetBulletFontFollowText();
        }

        /// <summary>
        /// <para>
        /// Get or set the bullet size that is to be used within a paragraph. This may be
        /// specified in three different ways, follows text size, percentage size and
        /// font points size:
        /// </para>
        /// <para>
        /// If value is instance of XDDFBulletSizeFollowText, then bullet size
        /// is text size; If given value is instance of XDDFBulletSizePercent, then
        /// bullet size is a percentage of the font size; If given value is instance
        /// of XDDFBulletSizePoints, then bullet size is specified in points.
        /// </para>
        /// </summary>

        public IXDDFBulletSize BulletSize
        {
            get
            {
                return FindDefinedParagraphProperty(
                    props => props.IsSetBuSzPct() || props.IsSetBuSzPts() || props.IsSetBuSzTx(),
                    props => new XDDFParagraphBulletProperties(props).BulletSize).OrElse(null);
            }
            set
            {
                if(value != null || _p.IsSetPPr())
                {
                    GetOrCreateBulletProperties().BulletSize = value;
                }
            }
        }


        public IXDDFBulletStyle BulletStyle
        {
            get
            {
                return FindDefinedParagraphProperty(
                props => props.IsSetBuAutoNum() || props.IsSetBuBlip() || props.IsSetBuChar() || props.IsSetBuNone(),
                props => new XDDFParagraphBulletProperties(props).BulletStyle).OrElse(null);
            }
            set
            {
                if(value != null || _p.IsSetPPr())
                {
                    GetOrCreateBulletProperties().BulletStyle = value;
                }
            }
        }

        public bool? HasEastAsianLineBreak
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.eaLnBrkSpecified,
                    props => props.eaLnBrk).OrElse(false);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetEastAsianLineBreak(value);
                }
            }
        }

        public bool? HasLatinLineBreak
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.latinLnBrkSpecified, props => props.latinLnBrk)
                    .OrElse(false);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetLatinLineBreak(value);
                }
            }
        }

        public bool? HasHangingPunctuation
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.hangingPunctSpecified, props => props.hangingPunct)
                    .OrElse(false);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetHangingPunctuation(value);
                }
            }
        }

        public bool? IsRightToLeft
        {
            get
            {
                return FindDefinedParagraphValueProperty(props => props.rtlSpecified, props => props.rtl).OrElse(false);
            }
            set
            {
                if(value.HasValue || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetRightToLeft(value);
                }
            }
        }

        public XDDFTabStop AddTabStop()
        {
            return GetOrCreateProperties().AddTabStop();
        }

        public XDDFTabStop InsertTabStop(int index)
        {
            return GetOrCreateProperties().InsertTabStop(index);
        }

        public void RemoveTabStop(int index)
        {
            if(_p.IsSetPPr())
            {
                GetProperties().RemoveTabStop(index);
            }
        }

        public XDDFTabStop GetTabStop(int index)
        {
            if(_p.IsSetPPr())
            {
                return GetProperties().GetTabStop(index);
            }
            else
            {
                return null;
            }
        }

        public List<XDDFTabStop> GetTabStops()
        {
            if(_p.IsSetPPr())
            {
                return GetProperties().GetTabStops();
            }
            else
            {
                return new List<XDDFTabStop>();
            }
        }

        public int CountTabStops()
        {
            if(_p.IsSetPPr())
            {
                return GetProperties().CountTabStops();
            }
            else
            {
                return 0;
            }
        }

        public XDDFParagraphBulletProperties GetOrCreateBulletProperties()
        {
            return GetOrCreateProperties().GetBulletProperties();
        }

        public XDDFParagraphBulletProperties BulletProperties
        {
            get
            {
                if(_p.IsSetPPr())
                {
                    return GetProperties().GetBulletProperties();
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <remarks>
        /// @since 4.0.1
        /// </remarks>
        public XDDFRunProperties AddDefaultRunProperties()
        {
            return GetOrCreateProperties().AddDefaultRunProperties();
        }

        public XDDFRunProperties DefaultRunProperties
        {
            get
            {
                if(_p.IsSetPPr())
                {
                    return GetProperties().GetDefaultRunProperties();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value != null || _p.IsSetPPr())
                {
                    GetOrCreateProperties().SetDefaultRunProperties(value);
                }
            }
        }

        public XDDFRunProperties AddAfterLastRunProperties()
        {
            if(!_p.IsSetEndParaRPr())
            {
                _p.AddNewEndParaRPr();
            }
            return AfterLastRunProperties;
        }

        public XDDFRunProperties AfterLastRunProperties
        {
            get
            {
                if(_p.IsSetEndParaRPr())
                {
                    return new XDDFRunProperties(_p.endParaRPr);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    if(_p.IsSetEndParaRPr())
                    {
                        _p.UnsetEndParaRPr();
                    }
                }
                else
                {
                    _p.endParaRPr = value.GetXmlObject();
                }
            }
        }


        private XDDFSpacing ExtractSpacing(CT_TextSpacing spacing)
        {
            if(spacing.IsSetSpcPct())
            {
                double scale = 1 - _parent.BodyProperties.AutoFit.LineSpaceReduction / 100_000.0;
                return new XDDFSpacingPercent(spacing, spacing.spcPct, scale);
            }
            else if(spacing.IsSetSpcPts())
            {
                return new XDDFSpacingPoints(spacing, spacing.spcPts);
            }
            return null;
        }

        private XDDFParagraphProperties GetProperties()
        {
            if(_properties == null)
            {
                _properties = new XDDFParagraphProperties(_p.pPr);
            }
            return _properties;
        }

        private XDDFParagraphProperties GetOrCreateProperties()
        {
            if(!_p.IsSetPPr())
            {
                _properties = new XDDFParagraphProperties(_p.AddNewPPr());
            }
            return GetProperties();
        }

        protected Option<R> FindDefinedParagraphProperty<R>(Func<CT_TextParagraphProperties, Boolean> isSet,
            Func<CT_TextParagraphProperties, R> getter) where R : class
        {
            if(_p.IsSetPPr())
            {
                int level = _p.pPr.lvlSpecified ? 1 + _p.pPr.lvl : 0;
                return FindDefinedParagraphProperty(isSet, getter, level);
            }
            else
            {
                return _parent.FindDefinedParagraphProperty(isSet, getter, 0);
            }
        }

        private Option<R> FindDefinedParagraphProperty<R>(Func<CT_TextParagraphProperties, Boolean> isSet,
            Func<CT_TextParagraphProperties, R> getter, int level) where R : class
        {
            CT_TextParagraphProperties props = _p.pPr;
            if(props != null && isSet.Invoke(props))
            {
                return Option<R>.Some(getter.Invoke(props));
            }
            else
            {
                return _parent.FindDefinedParagraphProperty(isSet, getter, level);
            }
        }

        internal Option<R> FindDefinedRunProperty<R>(Func<CT_TextCharacterProperties, Boolean> isSet,
            Func<CT_TextCharacterProperties, R> getter) where R : class
        {
            if(_p.IsSetPPr())
            {
                int level = _p.pPr.lvlSpecified ? 1 + _p.pPr.lvl : 0;
                return FindDefinedRunProperty(isSet, getter, level);
            }
            else
            {
                return _parent.FindDefinedRunProperty(isSet, getter, 0);
            }
        }

        private Option<R> FindDefinedRunProperty<R>(Func<CT_TextCharacterProperties, Boolean> isSet,
            Func<CT_TextCharacterProperties, R> getter, int level) where R : class
        {
            CT_TextCharacterProperties props = _p.pPr.IsSetDefRPr() ? _p.pPr.defRPr : null;
            if(props != null && isSet.Invoke(props))
            {
                return Option<R>.Some(getter.Invoke(props));
            }
            else
            {
                return _parent.FindDefinedRunProperty(isSet, getter, level);
            }
        }


        protected ValueOption<R> FindDefinedParagraphValueProperty<R>(Func<CT_TextParagraphProperties, Boolean> isSet,
            Func<CT_TextParagraphProperties, R> getter) where R : struct
        {
            if(_p.IsSetPPr())
            {
                int level = _p.pPr.lvlSpecified ? 1 + _p.pPr.lvl : 0;
                return FindDefinedParagraphValueProperty(isSet, getter, level);
            }
            else
            {
                return _parent.FindDefinedParagraphValueProperty(isSet, getter, 0);
            }
        }

        private ValueOption<R> FindDefinedParagraphValueProperty<R>(Func<CT_TextParagraphProperties, Boolean> isSet,
            Func<CT_TextParagraphProperties, R> getter, int level) where R : struct
        {
            CT_TextParagraphProperties props = _p.pPr;
            if(props != null && isSet.Invoke(props))
            {
                return ValueOption<R>.Some(getter.Invoke(props));
            }
            else
            {
                return _parent.FindDefinedParagraphValueProperty(isSet, getter, level);
            }
        }

        internal ValueOption<R> FindDefinedRunValueProperty<R>(Func<CT_TextCharacterProperties, Boolean> isSet,
            Func<CT_TextCharacterProperties, R> getter) where R : struct
        {
            if(_p.IsSetPPr())
            {
                int level = _p.pPr.lvlSpecified ? 1 + _p.pPr.lvl : 0;
                return FindDefinedRunValueProperty(isSet, getter, level);
            }
            else
            {
                return _parent.FindDefinedRunValueProperty(isSet, getter, 0);
            }
        }

        private ValueOption<R> FindDefinedRunValueProperty<R>(Func<CT_TextCharacterProperties, Boolean> isSet,
            Func<CT_TextCharacterProperties, R> getter, int level) where R : struct
        {
            CT_TextCharacterProperties props = _p.pPr.IsSetDefRPr() ? _p.pPr.defRPr : null;
            if(props != null && isSet.Invoke(props))
            {
                return ValueOption<R>.Some(getter.Invoke(props));
            }
            else
            {
                return _parent.FindDefinedRunValueProperty(isSet, getter, level);
            }
        }
    }
}
