using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;

namespace NPOI.XWPF.Usermodel
{
    public class XWPFOMathArg : MathContainer
    {
        public XWPFOMathArg(IOMathContainer c, IRunBody p) : base(c, p)
        {
        }
    }
}
