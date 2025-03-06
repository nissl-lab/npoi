using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SL.Draw.Geom
{
    public interface IAdjustableShape
    {
        GuideIf GetAdjustValue(string name);
    }
}
