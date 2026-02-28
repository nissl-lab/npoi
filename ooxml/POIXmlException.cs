using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI
{
    public class POIXMLException:Exception
    {
        public POIXMLException()
            : base()
        { }

        public POIXMLException(string msg)
            : base(msg)
        { }

        public POIXMLException(string msg,Exception ex)
            : base(msg,ex)
        { }

        public POIXMLException(Exception ex)
            : base(string.Empty,ex)
        { }
    }
}
