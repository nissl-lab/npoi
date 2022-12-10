using NPOI.Common.UserModel.Fonts;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSLF.UserModel
{
	public class HSLFFontInfoPredefined : FontInfo
	{
		public static readonly HSLFFontInfoPredefined ARIAL = new HSLFFontInfoPredefined("Arial", FontCharset.ANSI, FontPitchEnum.VARIABLE, FontFamilyEnum.FF_SWISS);
		public static readonly HSLFFontInfoPredefined TIMES_NEW_ROMAN = new HSLFFontInfoPredefined("Times New Roman", FontCharset.ANSI, FontPitchEnum.VARIABLE, FontFamilyEnum.FF_ROMAN);
		public static readonly HSLFFontInfoPredefined COURIER_NEW = new HSLFFontInfoPredefined("Courier New", FontCharset.ANSI, FontPitchEnum.FIXED, FontFamilyEnum.FF_MODERN);
		public static readonly HSLFFontInfoPredefined WINGDINGS = new HSLFFontInfoPredefined("Wingdings", FontCharset.SYMBOL, FontPitchEnum.VARIABLE, FontFamilyEnum.FF_DONTCARE);

		private String typeface;
		private FontCharset charset;
		private FontPitchEnum pitch;
		private FontFamilyEnum family;

		public HSLFFontInfoPredefined(String typeface, FontCharset charset, FontPitchEnum pitch, FontFamilyEnum family)
		{
			this.typeface = typeface;
			this.charset = charset;
			this.pitch = pitch;
			this.family = family;
		}

		//@Override
		public String GetTypeface()
		{
			return typeface;
		}

		//@Override
		public FontCharset GetCharset()
		{
			return charset;
		}

		//@Override
		public FontFamily GetFamily()
		{
			return FontFamily.ValueOf((int)family);
		}

		//@Override
		public FontPitch GetPitch()
		{
			return FontPitch.ValueOf((int)pitch);
		}

		public int GetIndex()
		{
			return 0;
		}

		public void SetIndex(int index)
		{
			throw new NotImplementedException();
		}

		public void SetTypeface(string typeface)
		{
			throw new NotImplementedException();
		}

		public void SetCharset(FontCharset charset)
		{
			throw new NotImplementedException();
		}

		public void SetFamily(FontFamily family)
		{
			throw new NotImplementedException();
		}

		public void SetPitch(FontPitch pitch)
		{
			throw new NotImplementedException();
		}

		public byte[] GetPanose()
		{
			return null;
		}

		public void SetPanose(byte[] panose)
		{
			throw new NotImplementedException();
		}

		public List<T> GetFacets<T>()
		{
			return new List<T>();
		}
	}
}
