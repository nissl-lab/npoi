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


using NPOI.Common.UserModel.Fonts;
using NPOI.HSLF.Exceptions;
using NPOI.HSLF.Model;
using NPOI.HSLF.Record;
using NPOI.SL.UserModel;
using NPOI.Util;
using System;
using static NPOI.HSLF.Model.TextPropCollection;

namespace NPOI.HSLF.UserModel
{
	public class HSLFTextRun: TextRun
	{
		/** The TextRun we belong to */
		private HSLFTextParagraph parentParagraph;
		private String _runText = "";
		/** Caches the font info objects until the text runs are attached to the container */
		private HSLFFontInfo[] cachedFontInfo;
		private HSLFHyperlink link;

		/**
		 * Our paragraph and character style.
		 * Note - we may share these styles with other RichTextRuns
		 */
		private TextPropCollection characterStyle = new TextPropCollection(1, TextPropType.character);

		/**
		 * Create a new wrapper around a rich text string
		 * @param parentParagraph the parent paragraph
		 */
		public HSLFTextRun(HSLFTextParagraph parentParagraph)
		{
			this.parentParagraph = parentParagraph;
		}

		public TextPropCollection GetCharacterStyle()
		{
			return characterStyle;
		}

		public void SetCharacterStyle(TextPropCollection characterStyle)
		{
			this.characterStyle = characterStyle.Copy();
			this.characterStyle.UpdateTextSize(_runText.Length);
		}

		/**
		 * Supply the SlideShow we belong to
		 */
		public void UpdateSheet()
		{
			if (cachedFontInfo != null)
			{
				foreach (FontGroupEnum tt in Enum.GetValues(typeof(FontGroupEnum)))
				{
					SetFontInfo(cachedFontInfo[(int)tt], tt);
				}
				cachedFontInfo = null;
			}
		}

		/**
		 * Get the length of the text
		 */
		public int GetLength()
		{
			return _runText.Length;
		}

		/**
		 * Fetch the text, in raw storage form
		 */
		//@Override
	public String GetRawText()
		{
			return _runText;
		}

		/**
		 * Change the text
		 */
		//@Override
	public void SetText(String text)
		{
			if (text == null)
			{
				throw new HSLFException("text must not be null");
			}
			String newText = HSLFTextParagraph.ToInternalString(text);
			if (!newText.Equals(_runText))
			{
				_runText = newText;
				if (HSLFSlideShow.GetLoadSavePhase() == HSLFSlideShow.LoadSavePhase.LOADED)
				{
					parentParagraph.SetDirty();
				}
			}
		}

		// --------------- Internal helpers on rich text properties -------

		/**
		 * Fetch the value of the given flag in the CharFlagsTextProp.
		 * Returns false if the CharFlagsTextProp isn't present, since the
		 *  text property won't be set if there's no CharFlagsTextProp.
		 */
		private bool IsCharFlagsTextPropVal(int index)
		{
			return GetFlag(index);
		}

		bool GetFlag(int index)
		{
			BitMaskTextProp prop = (characterStyle == null) ? null : characterStyle.FindByName<BitMaskTextProp>(CharFlagsTextProp.NAME);

			if (prop == null || !prop.GetSubPropMatches()[index])
			{
				prop = GetMasterProp<BitMaskTextProp>();
			}

			return prop != null && prop.GetSubValue(index);
		}

		private T GetMasterProp<T>() where T: TextProp
		{
			int txtype = parentParagraph.GetRunType();
			HSLFSheet sheet = parentParagraph.GetSheet();
			if (sheet == null)
			{
				//LOG.atError().log("Sheet is not available");
				return null;
			}

			HSLFMasterSheet master = sheet.GetMasterSheet();
			if (master == null)
			{
				//LOG.atWarn().log("MasterSheet is not available");
				return null;
			}

			String name = CharFlagsTextProp.NAME;
			TextPropCollection col = master.GetPropCollection(txtype, parentParagraph.GetIndentLevel(), name, true);
			return (T)((col == null) ? null : col.FindByName<TextProp>(name));
		}


		/**
		 * Set the value of the given flag in the CharFlagsTextProp, adding
		 *  it if required.
		 */
		private void SetCharFlagsTextPropVal(int index, bool value)
		{
			// TODO: check if paragraph/chars can be handled the same ...
			if (GetFlag(index) != value)
			{
				SetFlag(index, value);
				parentParagraph.SetDirty();
			}
		}

		/**
		 * Sets the value of the given Paragraph TextProp, add if required
		 * @param propName The name of the Paragraph TextProp
		 * @param val The value to set for the TextProp
		 */
		public void SetCharTextPropVal(String propName, int val)
		{
			GetTextParagraph().SetPropVal(characterStyle, propName, val);
		}


		// --------------- Friendly getters / setters on rich text properties -------

		//@Override
	public bool IsBold()
		{
			return IsCharFlagsTextPropVal(CharFlagsTextProp.BOLD_IDX);
		}

		//@Override
	public void SetBold(bool bold)
		{
			SetCharFlagsTextPropVal(CharFlagsTextProp.BOLD_IDX, bold);
		}

		//@Override
	public bool IsItalic()
		{
			return IsCharFlagsTextPropVal(CharFlagsTextProp.ITALIC_IDX);
		}

		//@Override
	public void SetItalic(bool italic)
		{
			SetCharFlagsTextPropVal(CharFlagsTextProp.ITALIC_IDX, italic);
		}

		//@Override
	public bool IsUnderlined()
		{
			return IsCharFlagsTextPropVal(CharFlagsTextProp.UNDERLINE_IDX);
		}

		//@Override
	public void SetUnderlined(bool underlined)
		{
			SetCharFlagsTextPropVal(CharFlagsTextProp.UNDERLINE_IDX, underlined);
		}

		/**
		 * Does the text have a shadow?
		 */
		public bool IsShadowed()
		{
			return IsCharFlagsTextPropVal(CharFlagsTextProp.SHADOW_IDX);
		}

		/**
		 * Does the text have a shadow?
		 */
		public void SetShadowed(bool flag)
		{
			SetCharFlagsTextPropVal(CharFlagsTextProp.SHADOW_IDX, flag);
		}

		/**
		 * Is this text embossed?
		 */
		public bool IsEmbossed()
		{
			return IsCharFlagsTextPropVal(CharFlagsTextProp.RELIEF_IDX);
		}

		/**
		 * Is this text embossed?
		 */
		public void SetEmbossed(bool flag)
		{
			SetCharFlagsTextPropVal(CharFlagsTextProp.RELIEF_IDX, flag);
		}

		//@Override
	public bool IsStrikethrough()
		{
			return IsCharFlagsTextPropVal(CharFlagsTextProp.STRIKETHROUGH_IDX);
		}

		//@Override
	public void SetStrikethrough(bool flag)
		{
			SetCharFlagsTextPropVal(CharFlagsTextProp.STRIKETHROUGH_IDX, flag);
		}

		/**
		 * Gets the subscript/superscript option
		 *
		 * @return the percentage of the font size. If the value is positive, it is superscript, otherwise it is subscript
		 */
		public int GetSuperscript()
		{
			TextProp tp = GetTextParagraph().GetPropVal<TextProp>(characterStyle, "superscript");
			return tp == null ? 0 : tp.GetValue();
		}

		/**
		 * Sets the subscript/superscript option
		 *
		 * @param val the percentage of the font size. If the value is positive, it is superscript, otherwise it is subscript
		 */
		public void SetSuperscript(int val)
		{
			SetCharTextPropVal("superscript", val);
		}

		//@Override
	public Double GetFontSize()
		{
			TextProp tp = GetTextParagraph().GetPropVal<TextProp>(characterStyle, "font.size");
			return tp == null ? 0 : (double)tp.GetValue();
		}


		//@Override
	public void SetFontSize(Double fontSize)
		{
			int iFontSize = (fontSize == null) ? 0 : (int)fontSize;
			SetCharTextPropVal("font.size", iFontSize);
		}

		/**
		 * Gets the font index
		 */
		public int GetFontIndex()
		{
			TextProp tp = GetTextParagraph().GetPropVal<TextProp>(characterStyle, "font.index");
			return tp == null ? -1 : tp.GetValue();
		}

		/**
		 * Sets the font index
		 */
		public void SetFontIndex(int idx)
		{
			SetCharTextPropVal("font.index", idx);
		}

		//@Override
	public void SetFontFamily(String typeface)
		{
			SetFontFamily(typeface, FontGroupEnum.LATIN);
		}

		//@Override
	public void SetFontFamily(String typeface, FontGroupEnum fontGroup)
		{
			SetFontInfo(new HSLFFontInfo(typeface), fontGroup);
		}

		//@Override
	public void SetFontInfo(FontInfo fontInfo, FontGroupEnum fontGroup)
		{
			FontGroupEnum fg = SafeFontGroup(fontGroup);

			HSLFSheet sheet = parentParagraph.GetSheet();
			//@SuppressWarnings("resource")

		HSLFSlideShow slideShow = (sheet == null) ? null : sheet.GetSlideShow();
			if (sheet == null || slideShow == null)
			{
				// we can't set font since slideshow is not assigned yet
				if (cachedFontInfo == null)
				{
					cachedFontInfo = new HSLFFontInfo[Enum.GetValues(typeof( FontGroup)).Length];
				}
				cachedFontInfo[(int)fg] = (fontInfo != null) ? new HSLFFontInfo(fontInfo) : null;
				return;
			}

			String propName;
			switch (fg)
			{
				default:
				case FontGroupEnum.LATIN:
					propName = "ansi.font.index";
					break;
				case FontGroupEnum.COMPLEX_SCRIPT:
				// TODO: implement TextCFException10 structure
				case FontGroupEnum.EAST_ASIAN:
					propName = "asian.font.index";
					break;
				case FontGroupEnum.SYMBOL:
					propName = "symbol.font.index";
					break;
			}


			// Get the index for this font, if it is not to be removed (typeface == null)
			Integer fontIdx = null;
			if (fontInfo != null)
			{
				fontIdx = slideShow.AddFont(fontInfo).getIndex();
			}


			SetCharTextPropVal("font.index", fontIdx);
			SetCharTextPropVal(propName, fontIdx);
		}

		//@Override
	public String GetFontFamily()
		{
			return GetFontFamily(null);
		}

		//@Override
	public String GetFontFamily(FontGroupEnum fontGroup)
		{
			HSLFFontInfo fi = GetFontInfo(fontGroup);
			return (fi != null) ? fi.GetTypeface() : null;
		}

		//@Override
	public HSLFFontInfo GetFontInfo(FontGroupEnum fontGroup)
		{
			FontGroupEnum fg = SafeFontGroup(fontGroup);

			HSLFSheet sheet = parentParagraph.GetSheet();
			//@SuppressWarnings("resource")

		HSLFSlideShow slideShow = (sheet == null) ? null : sheet.GetSlideShow();
			if (sheet == null || slideShow == null)
			{
				return (cachedFontInfo != null) ? cachedFontInfo[(int)fg] : null;
			}

			String propName;
			switch (fg)
			{
				default:
				case FontGroupEnum.LATIN:
					propName = "font.index,ansi.font.index";
					break;
				case FontGroupEnum.COMPLEX_SCRIPT:
				case FontGroupEnum.EAST_ASIAN:
					propName = "asian.font.index";
					break;
				case FontGroupEnum.SYMBOL:
					propName = "symbol.font.index";
					break;
			}

			TextProp tp = GetTextParagraph().GetPropVal<TextProp>(characterStyle, propName);
			return (tp != null) ? slideShow.GetFont(tp.GetValue()) : null;
		}

		/**
		 * @return font color as PaintStyle
		 */
		//@Override
	//public SolidPaint GetFontColor()
	//	{
	//		TextProp tp = getTextParagraph().getPropVal(characterStyle, "font.color");
	//		if (tp == null)
	//		{
	//			return null;
	//		}
	//		Color color = HSLFTextParagraph.getColorFromColorIndexStruct(tp.getValue(), parentParagraph.getSheet());
	//		return DrawPaint.createSolidPaint(color);
	//	}

		/**
		 * Sets color of the text, as a int bgr.
		 * (PowerPoint stores as BlueGreenRed, not the more
		 *  usual RedGreenBlue)
		 * @see Color
		 */
		//public void SetFontColor(int bgr)
		//{
		//	setCharTextPropVal("font.color", bgr);
		//}


		//@Override
	//public void SetFontColor(Color color)
	//	{
	//		SetFontColor(DrawPaint.createSolidPaint(color));
	//	}

		//@Override
	//public void SetFontColor(PaintStyle color)
	//	{
	//		if (!(color instanceof SolidPaint)) {
	//			throw new IllegalArgumentException("HSLF only supports solid paint");
	//		}
	//		// In PowerPont RGB bytes are swapped, as BGR
	//		SolidPaint sp = (SolidPaint)color;
	//		Color c = DrawPaint.applyColorTransform(sp.getSolidColor());
	//		int rgb = new Color(c.getBlue(), c.getGreen(), c.getRed(), 254).getRGB();
	//		SetFontColor(rgb);
	//	}

		private void SetFlag(int index, bool value)
		{
			BitMaskTextProp prop = characterStyle.AddWithName<BitMaskTextProp>(CharFlagsTextProp.NAME);
			prop.SetSubValue(value, index);
		}

		public HSLFTextParagraph GetTextParagraph()
		{
			return parentParagraph;
		}

		//@Override
	public TextCap GetTextCap()
		{
			return TextCap.NONE;
		}

		//@Override
	public bool IsSubscript()
		{
			return GetSuperscript() < 0;
		}

		//@Override
	public bool IsSuperscript()
		{
			return GetSuperscript() > 0;
		}

		//@Override
	public byte GetPitchAndFamily()
		{
			return 0;
		}

		/**
		 * Sets the hyperlink - used when parsing the document
		 *
		 * @param link the hyperlink
		 */
		/* package */
		public void SetHyperlink(HSLFHyperlink link)
		{
			this.link = link;
		}

		//@Override
	public HSLFHyperlink GetHyperlink()
		{
			return link;
		}

		//@Override
	public HSLFHyperlink CreateHyperlink()
		{
			if (link == null)
			{
				link = HSLFHyperlink.CreateHyperlink(this);
				parentParagraph.SetDirty();
			}
			return link;
		}

		//@Override
	public FieldType GetFieldType()
		{
			HSLFTextShape ts = GetTextParagraph().GetParentShape();
			Placeholder ph = ts.GetPlaceholder();

			if (ph != null)
			{
				switch (ph.nativeEnum)
				{
					case "SLIDE_NUMBER":
						return FieldType.SLIDE_NUMBER;
					case "DATETIME":
						return FieldType.DATE_TIME;
					default:
						break;
				}
			}

			Shape <object, object> ms = (ts.GetSheet() is MasterSheet) ? ts.GetMetroShape() : null;
			if (ms is TextShape) {
				return Stream.of((TextShape <?,?>)ms).
					flatMap(tsh-> ((List <? extends TextParagraph <?,?,? extends TextRun >>)tsh.GetTextParagraphs()).stream()).
					flatMap(tph->tph.getTextRuns().stream()).
					findFirst().
					map(TextRun::getFieldType).
					orElse(null);
			}

			return null;
		}

		private FontGroupEnum SafeFontGroup(FontGroupEnum fontGroup)
		{
			return (fontGroup != null) ? fontGroup : FontGroup.getFontGroupFirst(GetRawText());
		}

		//@Override
	public HSLFTextParagraph GetParagraph()
		{
			return parentParagraph;
		}
	}
}