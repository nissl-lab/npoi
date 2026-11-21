using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Common.UserModel.Fonts
{
    public interface FontFacet
    {
        int Weight { get; set; }
        bool IsItalic { get; set; }
        object GetFontData();
    }
}
