using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public interface ITableStyle
    {
        string Name { get; }
        int Index { get; }
        bool IsBuiltin { get; }
        DifferentialStyleProvider GetStyle(TableStyleType type);
    }
}
