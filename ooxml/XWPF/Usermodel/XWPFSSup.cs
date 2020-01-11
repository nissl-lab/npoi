using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;

namespace NPOI.XWPF.Usermodel
{
    /// <summary>
    /// Subscript Object
    /// his element specifies the subscript object sSub, which consists of a base e and a reduced-size scr placed below
    /// and to the right, as in Xn
    /// </summary>
    public class XWPFSSup : IRunBody
    {
        private CT_SSup ssup;
        private IRunBody parent;
        private readonly XWPFOMathArg e;
        private readonly XWPFOMathArg sup;

        public XWPFSSup(CT_SSup ssup, IRunBody p)
        {
            this.ssup = ssup;
            this.parent = p;

            if (ssup.e == null)
            {
                ssup.e = new CT_OMathArg();
            }
            this.e = new XWPFOMathArg(ssup.e, this);

            if (ssup.sup == null)
            {
                ssup.sup = new CT_OMathArg();
            }
            this.sup = new XWPFOMathArg(ssup.sup, this);
        }

        //<inheritdoc/>
        public XWPFDocument Document { get { return parent.Document; } }

        //<inheritdoc/>
        public POIXMLDocumentPart Part { get { return parent.Part; } }


        /// <summary>
        /// This tag, which is an abbreviation for “element”, serves several functions (18 total) including that of the base
        /// argument of a mathematical object or function, the elements in an array, and the elements in boxes.If all
        /// subelements are omitted, this element specifies the presence of an empty argument.
        /// </summary>
        public XWPFOMathArg Element { get { return e; } }

        /// <summary>
        /// This element specifies the Superscript of the Pre-Sub-Superscript object sPre
        /// </summary>
        public XWPFOMathArg Superscript { get { return sup; } }
    }
}
