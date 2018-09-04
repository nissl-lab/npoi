using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;

namespace NPOI.XWPF.Usermodel
{
    /// <summary>
    /// Accent
    /// This element specifies the accent function, consisting of a base 
    /// and a combining diacritical mark. If accPr is
    /// omitted, the default accent is U+0302 (COMBINING CIRCUMFLEX ACCENT). 
    /// </summary>
    public class XWPFAcc : IRunBody
    {
        private CT_Acc acc;
        private IRunBody parent;
        private readonly XWPFOMathArg e;

        public XWPFAcc(CT_Acc acc, IRunBody p)
        {
            this.acc = acc;
            this.parent = p;
            if(acc.e == null)
            {
                acc.e = new CT_OMathArg();
            }
            this.e = new XWPFOMathArg(acc.e, this);
            if (acc.accPr == null)
            {
                acc.accPr = new CT_AccPr();
            }
            if (acc.accPr.chr == null)
            {
                acc.accPr.chr = new CT_Char();
            }
        }

        /// <summary>
        /// Single char or UTF, like: &#771;
        /// </summary>
        public string AccPr {
            get
            {
                return acc.accPr.chr.val;
            }
            set
            {
                acc.accPr.chr.val = value;
            }
        }

        /// <summary>
        /// This tag, which is an abbreviation for “element”, serves several functions (18 total) including that of the base
        /// argument of a mathematical object or function, the elements in an array, and the elements in boxes.If all
        /// subelements are omitted, this element specifies the presence of an empty argument.
        /// </summary>
        public XWPFOMathArg Element { get { return e; } }

        //<inheritdoc/>
        public XWPFDocument Document { get { return parent.Document; } }

        //<inheritdoc/>
        public POIXMLDocumentPart Part { get { return parent.Part; } }
    }
}
