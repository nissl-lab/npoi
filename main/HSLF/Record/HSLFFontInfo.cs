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
using NPOI.Util;
using System;
using System.Collections.Generic;

namespace NPOI.HSLF.Record
{
	public class HSLFFontInfo: FontInfo
	{
		public enum FontRenderType
		{
			raster, device, truetype
		}

		/** A bit that specifies whether a subset of this font is embedded. */
		private static BitField FLAGS_EMBED_SUBSETTED      = BitFieldFactory.GetInstance(0x01);
    /** Bits that specifies whether the font is a raster,device or truetype font. */
    private static BitField FLAGS_RENDER_FONTTYPE      = BitFieldFactory.GetInstance(0x07);
    /** A bit that specifies whether font substitution logic is not applied for this font. */
    private static BitField FLAGS_NO_FONT_SUBSTITUTION = BitFieldFactory.GetInstance(0x08);
    
    private int index = -1;
		private String typeface = "undefined";
		private FontCharset charset = FontCharset.ANSI;
		private FontRenderType renderType = FontRenderType.truetype;
		private FontFamilyEnum family = FontFamilyEnum.FF_SWISS;
		private FontPitchEnum pitch = FontPitchEnum.VARIABLE;
		private bool isSubsetted;
		private bool isSubstitutable = true;
		private List<FontEmbeddedData> facets = new List<FontEmbeddedData>();
		private FontEntityAtom fontEntityAtom;

		/**
		 * Creates a new instance of HSLFFontInfo with more or sensible defaults.<p>
		 * 
		 * If you don't use default fonts (see {@link HSLFFontInfoPredefined}) then the results
		 * of the font substitution will be better, if you also specify the other properties.
		 * 
		 * @param typeface the font name
		 */
		public HSLFFontInfo(String typeface)
		{
			SetTypeface(typeface);
		}

		/**
		 * Creates a new instance of HSLFFontInfo and initialize it from the supplied font atom
		 */
		public HSLFFontInfo(FontEntityAtom fontAtom)
		{
			fontEntityAtom = fontAtom;
			SetIndex(fontAtom.GetFontIndex());
			SetTypeface(fontAtom.GetFontName());
			SetCharset((FontCharset)(fontAtom.GetCharSet()));
			// assumption: the render type is exclusive
			switch (FLAGS_RENDER_FONTTYPE.GetValue(fontAtom.GetFontType()))
			{
				case 1:
					SetRenderType(FontRenderType.raster);
					break;
				case 2:
					SetRenderType(FontRenderType.device);
					break;
				default:
				case 4:
					SetRenderType(FontRenderType.truetype);
					break;
			}

			byte pitchAndFamily = (byte)fontAtom.GetPitchAndFamily();
			SetPitch(FontPitch.ValueOfPitchFamily(pitchAndFamily));
			SetFamily(FontFamily.ValueOfPitchFamily(pitchAndFamily));
			SetEmbedSubsetted(FLAGS_EMBED_SUBSETTED.IsSet(fontAtom.GetFontFlags()));
			SetFontSubstitutable(!FLAGS_NO_FONT_SUBSTITUTION.IsSet(fontAtom.GetFontType()));
		}

		public HSLFFontInfo(FontInfo fontInfo)
		{
			// don't copy font index on copy constructor - it depends on the FontCollection this record is in
			SetTypeface(fontInfo.GetTypeface());
			SetCharset(fontInfo.GetCharset());
			SetFamily(fontInfo.GetFamily());
			SetPitch(fontInfo.GetPitch().GetNative());
			if (fontInfo is HSLFFontInfo) {
				HSLFFontInfo hFontInfo = (HSLFFontInfo)fontInfo;
				SetRenderType(hFontInfo.GetRenderType());
				SetEmbedSubsetted(hFontInfo.IsEmbedSubsetted());
				SetFontSubstitutable(hFontInfo.IsFontSubstitutable());
			}
		}

		//@Override
	public int GetIndex()
		{
			return index;
		}

		//@Override
	public void SetIndex(int index)
		{
			this.index = index;
		}

		//@Override
	public String GetTypeface()
		{
			return typeface;
		}

		//@Override
	public void SetTypeface(String typeface)
		{
			if (typeface == null || string.IsNullOrEmpty(typeface))
			{
				throw new InvalidOperationException("typeface can't be null nor empty");
			}
			this.typeface = typeface;
		}

		//@Override
	public void SetCharset(FontCharset charset)
		{
			this.charset = (charset == null) ? FontCharset.ANSI : charset;
		}

		//@Override
	public FontCharset GetCharset()
		{
			return charset;
		}

		//@Override
	public FontFamilyEnum GetFamily()
		{
			return family;
		}

		//@Override
	public void SetFamily(FontFamilyEnum family)
		{
			this.family = (family == null) ? FontFamilyEnum.FF_SWISS : family;
		}

		//@Override
	public FontPitchEnum GetPitch()
		{
			return pitch;
		}

		//@Override
	public void SetPitch(FontPitchEnum pitch)
		{
			this.pitch = (pitch == null) ? FontPitchEnum.VARIABLE : pitch;

		}

		public FontRenderType GetRenderType()
		{
			return renderType;
		}

		public void SetRenderType(FontRenderType renderType)
		{
			this.renderType = (renderType == null) ? FontRenderType.truetype : renderType;
		}

		public bool IsEmbedSubsetted()
		{
			return isSubsetted;
		}

		public void SetEmbedSubsetted(bool embedSubset)
		{
			this.isSubsetted = embedSubset;
		}

		public bool IsFontSubstitutable()
		{
			return this.isSubstitutable;
		}

		public void SetFontSubstitutable(bool isSubstitutable)
		{
			this.isSubstitutable = isSubstitutable;
		}

		public FontEntityAtom CreateRecord()
		{
			//assert(fontEntityAtom == null);

			FontEntityAtom fnt = new FontEntityAtom();
			fontEntityAtom = fnt;

			fnt.SetFontIndex(GetIndex() << 4);
			fnt.SetFontName(GetTypeface());
			fnt.SetCharSet((int)GetCharset());
			fnt.SetFontFlags((byte)(IsEmbedSubsetted() ? 1 : 0));

			int typeFlag;
			switch (renderType)
			{
				case FontRenderType.device:
					typeFlag = FLAGS_RENDER_FONTTYPE.SetValue(0, 1);
					break;
				case FontRenderType.raster:
					typeFlag = FLAGS_RENDER_FONTTYPE.SetValue(0, 2);
					break;
				default:
				case FontRenderType.truetype:
					typeFlag = FLAGS_RENDER_FONTTYPE.SetValue(0, 4);
					break;
			}
			typeFlag = FLAGS_NO_FONT_SUBSTITUTION.Setbool(typeFlag, IsFontSubstitutable());
			fnt.SetFontType(typeFlag);

			fnt.SetPitchAndFamily(FontPitch.GetNativeId(pitch, family));
			return fnt;
		}

		public void addFacet(FontEmbeddedData facet)
		{
			facets.add(facet);
		}

		@Override
	public List<FontEmbeddedData> getFacets()
		{
			return facets;
		}

		@Internal
	public FontEntityAtom getFontEntityAtom()
		{
			return fontEntityAtom;
		}
	}
}