using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OOXML.XSSF.UserModel
{
    public class DefaultIndexedColorMap:IIndexedColorMap
    {
        public byte[] GetRGB(int index)
        {
            return GetDefaultRGB(index);
        }
        public static byte[] GetDefaultRGB(int index)
        {
            HSSFColor hssfColor = HSSFColor.GetIndexHash()[index];
            if (hssfColor == null) return null;
            byte[] rgbShort = hssfColor.GetTriplet();
            return rgbShort;
        }
    }
}
