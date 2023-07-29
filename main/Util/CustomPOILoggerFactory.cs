using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Util
{
    /// <summary>
    /// Helper class to circumvent Type.GetType & Activator.CreateInstance under AOT
    /// </summary>
    public abstract class CustomPOILoggerFactory
    {
        public abstract POILogger Create(string name);
    }
}
