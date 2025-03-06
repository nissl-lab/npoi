using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Common.UserModel.Fonts
{
    public class FontPitch
    {
        /**
       * The default pitch, which is implementation-dependent.
       */
        public static readonly FontPitch DEFAULT = new FontPitch(0x00);
        /**
         * A fixed pitch, which means that all the characters in the font occupy the same
         * width when output in a string.
         */
        public static readonly FontPitch FIXED = new FontPitch(0x01);
        /**
         * A variable pitch, which means that the characters in the font occupy widths
         * that are proportional to the actual widths of the glyphs when output in a string. For example,
         * the "i" and space characters usually have much smaller widths than a "W" or "O" character.
         */
        public static readonly FontPitch VARIABLE= new FontPitch(0x02);

        private int nativeId;
        protected FontPitch(int nativeId)
        {
            this.nativeId= nativeId;
        }
        public int NativeId {
            get {
                return nativeId;
            }
        }
        public static FontPitch ValueOf(int flag)
        {
            switch(flag)
            {
                case 1:
                    return FIXED;
                case 2:
                    return VARIABLE;
                default:
                    return DEFAULT;
            }
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
            return (byte) (pitch.NativeId | (family.NativeId << 4));
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
