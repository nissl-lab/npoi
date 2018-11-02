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
    public class XWPFSSub : IRunBody
    {
        private CT_SSub ssub;
        private IRunBody parent;
        private readonly XWPFOMathArg e;
        private readonly XWPFOMathArg sub;

        public XWPFSSub(CT_SSub ssub, IRunBody p)
        {
            this.ssub = ssub;
            this.parent = p;

            if (ssub.e == null)
            {
                ssub.e = new CT_OMathArg();
            }
            this.e = new XWPFOMathArg(ssub.e, this);

            if (ssub.sub == null)
            {
                ssub.sub = new CT_OMathArg();
            }
            this.sub = new XWPFOMathArg(ssub.sub, this);
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
        /// This element specifies the subscript of the Pre-Sub-Superscript object sPre
        /// </summary>
        public XWPFOMathArg Subscript { get { return sub; } }
    }
}
