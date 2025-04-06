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
    using NPOI.Common.UserModel.Fonts;
    using NPOI.OOXML;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using NPOI.XDDF.UserModel;
    using NPOI.XSSF.UserModel;
    using NPOI.OpenXmlFormats.Dml;
    using System.ComponentModel;
    using System.Globalization;
    using NPOI.Util.Optional;

    public class XDDFTextRun
    {
        private XDDFTextParagraph _parent;
        private XDDFRunProperties _properties;
        private CT_TextLineBreak _tlb;
        private CT_TextField _tf;
        private CT_RegularTextRun _rtr;
        internal XDDFTextRun(CT_TextLineBreak run, XDDFTextParagraph parent)
        {
            this._tlb = run;
            this._parent = parent;
        }
        internal XDDFTextRun(CT_TextField run, XDDFTextParagraph parent)
        {
            this._tf = run;
            this._parent = parent;
        }
        internal XDDFTextRun(CT_RegularTextRun run, XDDFTextParagraph parent)
        {
            this._rtr = run;
            this._parent = parent;
        }

        public XDDFTextParagraph GetParentParagraph()
        {
            return _parent;
        }

        public bool IsLineBreak
        {
            get
            {
                return _tlb != null;
            }
        }

        public bool IsField
        {
            get
            {
                return _tf != null;
            }
        }

        public bool IsRegularRun
        {
            get
            {
                return _rtr != null;
            }
        }

        public string Text
        {
            get
            {
                if(IsLineBreak)
                {
                    return "\n";
                }
                else if(IsField)
                {
                    return _tf.t;
                }
                else
                {
                    return _rtr.t;
                }
            }
            set
            {
                if(IsField)
                {
                    _tf.t = value;
                }
                else if(IsRegularRun)
                {
                    _rtr.t = value;
                }
            }
        }

        public bool? Dirty
        {
            get => GetDirty();
            set => SetDirty(value);
        }
        private void SetDirty(bool? dirty)
        {
            GetOrCreateProperties().SetDirty(dirty);
        }

        private bool? GetDirty()
        {
            return FindDefinedValueProperty(props => props.dirtySpecified, props => props.dirty)
                .OrElse(null);
        }

        public bool? SpellError
        {
            get => GetSpellError();
            set => SetSpellError(value);
        }
        private void SetSpellError(bool? error)
        {
            GetOrCreateProperties().SetSpellError(error);
        }

        private Boolean? GetSpellError()
        {
            return FindDefinedValueProperty(props => props.errSpecified, props => props.err)
                .OrElse(null);
        }

        public bool? NoProof
        {
            get => GetNoProof();
            set => SetNoProof(value);
        }
        private void SetNoProof(bool? noproof)
        {
            GetOrCreateProperties().SetNoProof(noproof);
        }

        private Boolean? GetNoProof()
        {
            return FindDefinedValueProperty(props => props.noProofSpecified, props => props.noProof)
                .OrElse(null);
        }

        public bool? NormalizeHeights
        {
            get => GetNormalizeHeights();
            set => SetNormalizeHeights(value);
        }

        private void SetNormalizeHeights(bool? normalize)
        {
            GetOrCreateProperties().SetNormalizeHeights(normalize);
        }

        private Boolean? GetNormalizeHeights()
        {
            return FindDefinedValueProperty(props => props.noProofSpecified, props => props.normalizeH)
                .OrElse(null);
        }

        [Obsolete("use property IsKumimoji")]
        public void SetKumimoji(Boolean kumimoji)
        {
            GetOrCreateProperties().SetKumimoji(kumimoji);
        }

        public bool IsKumimoji
        {
            get
            {
                return FindDefinedValueProperty(props => props.kumimojiSpecified, props => props.kumimoji)
                    .OrElse(false);
            }
            set
            {
                GetOrCreateProperties().SetKumimoji(value);
            }
        }

        /// <summary>
        /// Specifies whether this run of text will be formatted as bold text.
        /// </summary>
        /// <param name="bold">
        /// whether this run of text will be formatted as bold text.
        /// </param>
        [Obsolete("use property IsBold")]
        public void SetBold(Boolean bold)
        {
            GetOrCreateProperties().SetBold(bold);
        }

        /// <summary>
        /// </summary>
        /// <returns>whether this run of text is formatted as bold text.</returns>
        public bool IsBold
        {
            get
            {
                return FindDefinedValueProperty(props => props.bSpecified, props => props.b)
                .OrElse(false);
            }
            set
            {
                GetOrCreateProperties().SetBold(value);
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="italic">
        /// whether this run of text is formatted as italic text.
        /// </param>
        [Obsolete("use property IsItalic")]
        public void SetItalic(Boolean italic)
        {
            GetOrCreateProperties().SetItalic(italic);
        }

        /// <summary>
        /// </summary>
        /// <returns>whether this run of text is formatted as italic text.</returns>
        public bool IsItalic
        {
            get
            {
                return FindDefinedValueProperty(props => props.iSpecified, props => props.i)
                .OrElse(false);
            }
            set
            {
                GetOrCreateProperties().SetItalic(value);
            }
        }

        public StrikeType? StrikeThrough
        {
            get => GetStrikeThrough();
            set => SetStrikeThrough(value);
        }

        /// <summary>
        /// </summary>
        /// <param name="strike">
        /// which strike style this run of text is formatted with.
        /// </param>
        private void SetStrikeThrough(StrikeType? strike)
        {
            GetOrCreateProperties().SetStrikeThrough(strike);
        }

        /// <summary>
        /// </summary>
        /// <returns>whether this run of text is formatted as striked text.</returns>
        public bool IsStrikeThrough
        {
            get
            {
                return FindDefinedValueProperty(props => props.IsSetStrike(), props => props.strike)
                    .MapValue(strike => strike != ST_TextStrikeType.noStrike)
                    .OrElse(false);
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>which strike style this run of text is formatted with.</returns>
        private StrikeType? GetStrikeThrough()
        {
            return FindDefinedValueProperty(props => props.IsSetStrike(), props => props.strike)
                .MapValue(strike => StrikeTypeExtensions.ValueOf(strike))
                .OrElse(null);
        }

        /// <summary>
        /// which underline style this run of text is formatted with.
        /// </summary>
        public UnderlineType? Underline
        {
            get => GetUnderline();
            set => SetUnderline(value);
        }

        /// <summary>
        /// </summary>
        /// <param name="underline">
        /// which underline style this run of text is formatted with.
        /// </param>
        private void SetUnderline(UnderlineType? underline)
        {
            GetOrCreateProperties().SetUnderline(underline);
        }

        /// <summary>
        /// </summary>
        /// <returns>whether this run of text is formatted as underlined text.</returns>
        public bool IsUnderline
        {
            get
            {
                return FindDefinedValueProperty(props => props.IsSetU(), props => props.u)
                    .MapValue(underline => underline != ST_TextUnderlineType.none)
                    .OrElse(false);
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>which underline style this run of text is formatted with.</returns>
        private UnderlineType? GetUnderline()
        {
            return FindDefinedValueProperty(props => props.IsSetU(), props => props.u)
                .MapValue(underline => UnderlineTypeExtensions.ValueOf(underline))
                .OrElse(null);
        }

        /// <summary>
        /// </summary>
        /// <param name="caps">
        /// which caps style this run of text is formatted with.
        /// </param>
        private void SetCapitals(CapsType? caps)
        {
            GetOrCreateProperties().SetCapitals(caps);
        }

        /// <summary>
        /// </summary>
        /// <returns>whether this run of text is formatted as capitalized text.</returns>
        public bool IsCapitals
        {
            get
            {
                return FindDefinedValueProperty(props => props.IsSetCap(), props => props.cap)
                .MapValue(caps => caps != ST_TextCapsType.none)
                .OrElse(false);
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>which caps style this run of text is formatted with.</returns>
        public CapsType? Capitals
        {
            get => GetCapitals();
            set => SetCapitals(value);
        }
        private CapsType? GetCapitals()
        {
            return FindDefinedValueProperty(props => props.IsSetCap(), props => props.cap)
                .MapValue(caps => CapsTypeExtensions.ValueOf(caps))
                .OrElse(null);
        }

        /// <summary>
        /// </summary>
        /// <returns>whether a run of text will be formatted as a subscript text.
        /// Default is false.
        /// </returns>
        public bool IsSubscript
        {
            get
            {
                return FindDefinedValueProperty(props => props.IsSetBaseline(), props => props.baseline)
                .MapValue(baseline => baseline < 0)
                .OrElse(false);
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>whether a run of text will be formatted as a superscript text.
        /// Default is false.
        /// </returns>
        public bool IsSuperscript
        {
            get
            {
                return FindDefinedValueProperty(props => props.IsSetBaseline(), props => props.baseline)
                .MapValue(baseline => baseline > 0)
                .OrElse(false);
            }
        }

        /// <summary>
        /// <para>
        ///  Set the baseline for both the superscript and subscript fonts.
        /// </para>
        /// <para>
        ///     The size is specified using a percentage.
        ///     Positive values indicate superscript, negative values indicate subscript.
        /// </para>
        /// </summary>
        /// <param name="offset"></param>
        public void SetBaseline(Double? offset)
        {
            if(offset == null)
            {
                GetOrCreateProperties().SetBaseline(null);
            }
            else
            {
                GetOrCreateProperties().SetBaseline((int) (offset * 1000));
            }
        }

        /// <summary>
        /// <para>
        /// Set whether the text in this run is formatted as superscript.
        /// </para>
        /// <para>
        /// The size is specified using a percentage.
        /// </para>
        /// </summary>
        /// <param name="offset"></param>
        public void SetSuperscript(Double? offset)
        {
            SetBaseline(offset.HasValue ? Math.Abs(offset.Value) : null);
        }

        /// <summary>
        /// <para>
        /// Set whether the text in this run is formatted as subscript.
        /// </para>
        /// <para>
        /// The size is specified using a percentage.
        /// </para>
        /// </summary>
        /// <param name="offset"></param>
        public void SetSubscript(Double? offset)
        {
            SetBaseline(offset.HasValue ? -Math.Abs(offset.Value) : null);
        }

        public void SetFillProperties(IXDDFFillProperties properties)
        {
            GetOrCreateProperties().SetFillProperties(properties);
        }

        public XDDFColor FontColor
        {
            get => GetFontColor();
            set => SetFontColor(value);
        }

        private void SetFontColor(XDDFColor color)
        {
            XDDFSolidFillProperties props = new XDDFSolidFillProperties();
            props.Color = color;
            SetFillProperties(props);
        }

        private XDDFColor GetFontColor()
        {
            XDDFSolidFillProperties solid = FindDefinedProperty(props => props.IsSetSolidFill(), props => props.solidFill)
            .Map(props => new XDDFSolidFillProperties(props))
            .OrElse(new XDDFSolidFillProperties());
            return solid.Color;
        }

        /// <summary>
        /// <em>Note</em>: In order to Get fonts to unset the property for a given font family use
        /// <see cref="XDDFFont.unsetFontForGroup(FontGroup)" />
        /// </summary>
        /// <param name="fonts">
        /// to Set or unset on the run.
        /// </param>
        public void SetFonts(XDDFFont[] fonts)
        {
            GetOrCreateProperties().SetFonts(fonts);
        }

        public XDDFFont[] GetFonts()
        {
            LinkedList<XDDFFont> list = new LinkedList<XDDFFont>();

            FindDefinedProperty(props => props.IsSetCs(), props => props.cs)
                .Map(font => new XDDFFont(FontGroup.COMPLEX_SCRIPT, font))
                .IfPresent(font => list.AddLast(font));
            FindDefinedProperty(props => props.IsSetEa(), props => props.ea)
                .Map(font => new XDDFFont(FontGroup.EAST_ASIAN, font))
                .IfPresent(font => list.AddLast(font));
            FindDefinedProperty(props => props.IsSetLatin(), props => props.latin)
                .Map(font => new XDDFFont(FontGroup.LATIN, font))
                .IfPresent(font => list.AddLast(font));
            FindDefinedProperty(props => props.IsSetSym(), props => props.sym)
                .Map(font => new XDDFFont(FontGroup.SYMBOL, font))
                .IfPresent(font => list.AddLast(font));

            return [.. list];
        }

        /// <summary>
        /// </summary>
        /// <param name="size">
        /// font size in points. The value <c>null</c> unsets the
        /// size for this run.
        /// <dl>
        /// <dt>Minimum inclusive =</dt>
        /// <dd>1</dd>
        /// <dt>Maximum inclusive =</dt>
        /// <dd>400</dd></dt>
        /// </param>
        ///
        public double FontSize
        {
            get => GetFontSize();
            set => SetFontSize(value);
        }
        private void SetFontSize(Double size)
        {
            GetOrCreateProperties().SetFontSize(size);
        }

        private Double GetFontSize()
        {
            int size = FindDefinedValueProperty(props => props.IsSetSz(), props => props.sz)
            .OrElse(100 * XSSFFont.DEFAULT_FONT_SIZE); // default font size
            double scale = _parent.GetParentBody().BodyProperties.AutoFit.FontScale / 10_000_000.0;
            return size * scale;
        }

        /// <summary>
        /// <para>
        /// Set the kerning of characters within a text run.
        /// </para>
        /// <para>
        /// The value <c>null</c> unsets the kerning for this run.
        /// </para>
        /// </summary>
        /// <param name="kerning">
        /// character kerning in points.
        /// <dl>
        /// <dt>Minimum inclusive =</dt>
        /// <dd>0</dd>
        /// <dt>Maximum inclusive =</dt>
        /// <dd>4000</dd></dt>
        /// </param>
        private void SetCharacterKerning(Double? kerning)
        {
            GetOrCreateProperties().SetCharacterKerning(kerning);
        }

        /// <summary>
        /// </summary>
        /// <returns>the kerning of characters within a text run,
        /// If this attribute is omitted then returns <c>null</c>.
        /// </returns>
        private Double? GetCharacterKerning()
        {
            return FindDefinedValueProperty(props => props.kernSpecified, props => props.kern)
                .MapValue(kerning => 0.01 * kerning)
                .OrElse(null);
        }

        /// <summary>
        /// <para>
        /// Set the kerning of characters within a text run.
        /// </para>
        /// <para>
        /// The value <c>null</c> unsets the kerning for this run.
        /// </para>

        /// character kerning in points.
        /// <dl>
        /// <dt>Minimum inclusive =</dt>
        /// <dd>0</dd>
        /// <dt>Maximum inclusive =</dt>
        /// <dd>4000</dd></dt>
        /// </summary>
        public double? CharacterKerning
        {
            get => GetCharacterKerning();
            set => SetCharacterKerning(value);
        }

        /// <summary>
        /// <para>
        /// Set the spacing between characters within a text run.
        /// </para>
        /// <para>
        /// The spacing is specified in points. Positive values will cause the text to expand,
        /// negative values to condense.
        /// </para>
        /// <para>
        /// </para>
        /// <para>
        /// The value <c>null</c> unsets the spacing for this run.
        /// </para>
        /// </summary>
        /// <param name="spacing">
        /// character spacing in points.
        /// <dl>
        /// <dt>Minimum inclusive =</dt>
        /// <dd>-4000</dd>
        /// <dt>Maximum inclusive =</dt>
        /// <dd>4000</dd></dt>
        /// </param>
        private void SetCharacterSpacing(Double? spacing)
        {
            GetOrCreateProperties().SetCharacterSpacing(spacing);
        }

        /// <summary>
        /// </summary>
        /// <returns>the spacing between characters within a text run,
        /// If this attribute is omitted then returns <c>null</c>.
        /// </returns>
        private Double? GetCharacterSpacing()
        {
            return FindDefinedValueProperty(props => props.IsSetSpc(), props => props.spc)
                .MapValue(spacing => 0.01 * spacing)
                .OrElse(null);
        }

        /// <summary>
        /// <para>
        /// Set the spacing between characters within a text run.
        /// </para>
        /// <para>
        /// The spacing is specified in points. Positive values will cause the text to expand,
        /// negative values to condense.
        /// </para>
        /// <para>
        /// </para>
        /// <para>
        /// The value <c>null</c> unsets the spacing for this run.
        /// </para>
        /// character spacing in points.
        /// <dl>
        /// <dt>Minimum inclusive =</dt>
        /// <dd>-4000</dd>
        /// <dt>Maximum inclusive =</dt>
        /// <dd>4000</dd></dt>
        /// </summary>
        public double? CharacterSpacing
        {
            get => GetCharacterSpacing();
            set => SetCharacterSpacing(value);
        }

        public string Bookmark
        {
            get => GetBookmark();
            set => SetBookmark(value);
        }

        private void SetBookmark(string bookmark)
        {
            GetOrCreateProperties().SetBookmark(bookmark);
        }

        private string GetBookmark()
        {
            return FindDefinedProperty(props => props.IsSetBmk(), props => props.bmk).OrElse(null);
        }

        public XDDFHyperlink LinkToExternal(string url, PackagePart localPart, POIXMLRelation relation)
        {
            PackageRelationship rel = localPart.AddExternalRelationship(url, relation.Relation);
            XDDFHyperlink link = new XDDFHyperlink(rel.Id);
            GetOrCreateProperties().SetHyperlink(link);
            return link;
        }

        public XDDFHyperlink LinkToAction(string action)
        {
            XDDFHyperlink link = new XDDFHyperlink("", action);
            GetOrCreateProperties().SetHyperlink(link);
            return link;
        }

        public XDDFHyperlink LinkToInternal(string action, PackagePart localPart, POIXMLRelation relation, PackagePartName target)
        {
            PackageRelationship rel = localPart.AddRelationship(target, TargetMode.Internal, relation.Relation);
            XDDFHyperlink link = new XDDFHyperlink(rel.Id, action);
            GetOrCreateProperties().SetHyperlink(link);
            return link;
        }

        public XDDFHyperlink GetHyperlink()
        {
            return FindDefinedProperty(props => props.IsSetHlinkClick(), props => props.hlinkClick)
                .Map(link => new XDDFHyperlink(link))
                .OrElse(null);
        }

        public XDDFHyperlink CreateMouseOver(string action)
        {
            XDDFHyperlink link = new XDDFHyperlink("", action);
            GetOrCreateProperties().SetMouseOver(link);
            return link;
        }

        public XDDFHyperlink GetMouseOver()
        {
            return FindDefinedProperty(props => props.IsSetHlinkMouseOver(), props => props.hlinkMouseOver)
                .Map(link => new XDDFHyperlink(link))
                .OrElse(null);
        }

        public CultureInfo Language
        {
            get => GetLanguage();
            set => SetLanguage(value);
        }

        private void SetLanguage(CultureInfo lang)
        {
            GetOrCreateProperties().SetLanguage(lang);
        }

        private CultureInfo GetLanguage()
        {
            return FindDefinedProperty(props => props.IsSetLang(), props => props.lang)
                .Map(lang => CultureInfo.GetCultureInfo(lang))
                .OrElse(null);
        }

        public CultureInfo AlternativeLanguage
        {
            get => GetAlternativeLanguage();
            set => SetAlternativeLanguage(value);
        }

        private void SetAlternativeLanguage(CultureInfo lang)
        {
            GetOrCreateProperties().SetAlternativeLanguage(lang);
        }

        private CultureInfo GetAlternativeLanguage()
        {
            return FindDefinedProperty(props => props.IsSetAltLang(), props => props.altLang)
                .Map(lang => CultureInfo.GetCultureInfo(lang))
                .OrElse(null);
        }

        public XDDFColor Highlight
        {
            get => GetHighlight();
            set => SetHighlight(value);
        }

        private void SetHighlight(XDDFColor color)
        {
            GetOrCreateProperties().SetHighlight(color);
        }

        private XDDFColor GetHighlight()
        {
            return FindDefinedProperty(props => props.IsSetHighlight(), props => props.highlight)
                .Map(color => XDDFColor.ForColorContainer(color))
                .OrElse(null);
        }

        public XDDFLineProperties LineProperties
        {
            get => GetLineProperties();
            set => SetLineProperties(value);
        }

        private void SetLineProperties(XDDFLineProperties properties)
        {
            GetOrCreateProperties().SetLineProperties(properties);
        }

        private XDDFLineProperties GetLineProperties()
        {
            return FindDefinedProperty(props => props.IsSetLn(), props => props.ln)
                .Map(props => new XDDFLineProperties(props))
                .OrElse(null);
        }

        private Option<R> FindDefinedProperty<R>(Func<CT_TextCharacterProperties, Boolean> isSet,
            Func<CT_TextCharacterProperties, R> getter) where R : class
        {
            CT_TextCharacterProperties props = GetProperties();
            if(props != null && isSet.Invoke(props))
            {
                return Option<R>.Some(getter.Invoke(props));
            }
            else
            {
                return _parent.FindDefinedRunProperty(isSet, getter);
            }
        }


        private ValueOption<R> FindDefinedValueProperty<R>(Func<CT_TextCharacterProperties, Boolean> isSet,
            Func<CT_TextCharacterProperties, R> getter) where R : struct
        {
            CT_TextCharacterProperties props = GetProperties();
            if(props != null && isSet.Invoke(props))
            {
                return ValueOption<R>.Some(getter.Invoke(props));
            }
            else
            {
                return _parent.FindDefinedRunValueProperty(isSet, getter);
            }
        }

        internal CT_TextCharacterProperties GetProperties()
        {
            if(IsLineBreak && _tlb.IsSetRPr())
            {
                return _tlb.rPr;
            }
            else if(IsField && _tf.IsSetRPr())
            {
                return _tf.rPr;
            }
            else if(IsRegularRun && _rtr.IsSetRPr())
            {
                return _rtr.rPr;
            }
            return null;
        }

        private XDDFRunProperties GetOrCreateProperties()
        {
            if(_properties == null)
            {
                if(IsLineBreak)
                {
                    _properties = new XDDFRunProperties(_tlb.IsSetRPr() ? _tlb.rPr : _tlb.AddNewRPr());
                }
                else if(IsField)
                {
                    _properties = new XDDFRunProperties(_tf.IsSetRPr() ? _tf.rPr : _tf.AddNewRPr());
                }
                else if(IsRegularRun)
                {
                    _properties = new XDDFRunProperties(_rtr.IsSetRPr() ? _rtr.rPr : _rtr.AddNewRPr());
                }
            }
            return _properties;
        }
    }
}