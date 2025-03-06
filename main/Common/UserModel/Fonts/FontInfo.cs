using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Common.UserModel.Fonts
{
    public interface FontInfo
    {
        int Index { get; set; }
        string Typeface { get; set; }
        FontCharset Charset { get; set; }
        FontFamily Family { get; set; }

        byte[] Panose { get; set; }

        List<FontFacet> GetFacets();
    }
}
