using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;

namespace NPOI.XWPF.Usermodel
{
    /// <summary>
    /// Radical Object
    /// This element specifies the radical object, consisting of a radical, a base e, and an optional degree deg. [Example:
    /// Example of rad are √𝑥
    /// </summary>
    public class XWPFRad : IRunBody
    {
        private CT_Rad rad;
        private IRunBody parent;
        private XWPFOMathArg deg;
        private XWPFOMathArg e;

        public XWPFRad(CT_Rad rad, IRunBody p)
        {
            this.rad = rad;
            this.parent = p;

            //TODO: Implement public radPr
            if (rad.radPr == null)
            {
                rad.radPr = new CT_RadPr();
            }

            if (rad.deg == null)
            {
                rad.deg = new CT_OMathArg();
            }
            this.deg = new XWPFOMathArg(rad.deg, this);

            if (rad.e == null)
            {
                rad.e = new CT_OMathArg();
            }

            this.e = new XWPFOMathArg(rad.e, this);
        }

        //<inheritdoc/>
        public XWPFDocument Document { get { return parent.Document; } }

        //<inheritdoc/>
        public POIXMLDocumentPart Part { get { return parent.Part; } }

        /// <summary>
        /// This element specifies the degree in the mathematical radical. This element is optional. When omitted, the
        /// square root function, as in √𝑥, is assumed.
        /// </summary>
        public XWPFOMathArg Degree { get { return deg; } }

        /// <summary>
        /// Radical expression element.
        /// </summary>
        public XWPFOMathArg Element { get { return e; } }
    }
}
