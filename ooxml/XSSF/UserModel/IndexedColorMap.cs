using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OOXML.XSSF.UserModel
{
    public interface IIndexedColorMap
    {
        byte[] GetRGB(int index);
    }
}
