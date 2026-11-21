using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SL.Draw.Geom
{
    public interface AdjustPointIf
    {
        string X { get; set; }
        bool IsSetX { get; }
        string Y { get; set; }
        bool IsSetY { get; }
    }
}
