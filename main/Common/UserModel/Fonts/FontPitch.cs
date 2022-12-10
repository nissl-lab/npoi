using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace NPOI.Common.UserModel.Fonts
{
	public class FontPitch
	{
		private int nativeId;
		private FontPitchEnum native;

	    public FontPitch(int nativeId)
		{
			this.nativeId = nativeId;
			this.native = (FontPitchEnum)nativeId;
		}
	
	    public int GetNativeId()
		{
			return nativeId;
		}
	
	    public static FontPitch ValueOf(int flag)
		{
			foreach (var item in Enum.GetValues(typeof(FontPitchEnum)))
			{
				if (flag == (int)item)
				{
					return new FontPitch(flag);
				}
			}
			return null;
		}
	    
	    /**
	     * Combine pitch and family to native id
	     * 
	     * @see <a href="https://msdn.microsoft.com/en-us/library/dd145037.aspx">LOGFONT structure</a>
	     *
	     * @param pitch The pitch-value, cannot be null
	     * @param family The family-value, cannot be null
	     *
	     * @return The resulting combined byte-value with pitch and family encoded into one byte
	     */
	    public static byte GetNativeId(FontPitch pitch, FontFamily family)
		{
			return (byte)(pitch.GetNativeId() | (family.GetFlag() << 4));
		}
	
	    /**
	     * Get FontPitch from native id
	     *
	     * @param pitchAndFamily The combined byte value for pitch and family
	     *
	     * @return The resulting FontPitch enumeration value
	     */
	    public static FontPitch ValueOfPitchFamily(byte pitchAndFamily)
		{
			return ValueOf(pitchAndFamily & 0x3);
		}
	}
}
