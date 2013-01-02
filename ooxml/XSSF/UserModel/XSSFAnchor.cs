using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.XSSF.UserModel
{
    /**
     * An anchor is what specifics the position of a shape within a client object
     * or within another containing shape.
     *
     * @author Yegor Kozlov
     */
    public abstract class XSSFAnchor
    {
        public abstract int Dx1 { get; set; }
        public abstract int Dy1 { get; set; }
        public abstract int Dy2 { get; set; }
        public abstract int Dx2 { get; set; }

    }
}
