using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;

namespace NPOI.XWPF.Usermodel
{
    /// <summary>
    /// Fraction Object
    /// This element specifies the fraction object, consisting of a numerator and denominator separated by a fraction
    /// bar.The fraction bar can be horizontal or diagonal, depending on the fraction properties.The fraction object is
    /// also used to represent the stack function, which places one element above another, with no fraction bar.
    /// </summary>
    public class XWPFF : IRunBody
    {
        private CT_F f;
        private IRunBody parent;
        private XWPFOMathArg num;
        private XWPFOMathArg den;

        public XWPFF(CT_F f, IRunBody p)
        {
            this.f = f;
            this.parent = p;

            if (f.fPr == null)
            {
                f.fPr = new CT_FPr();
            }

            if (f.fPr.type == null)
            {
                f.fPr.type = new CT_FType();
            }

            if (f.num == null)
            {
                f.num = new CT_OMathArg();
            }
            this.num = new XWPFOMathArg(f.num, this);

            if (f.den == null)
            {
                f.den = new CT_OMathArg();
            }
            this.den = new XWPFOMathArg(f.den, this);
        }

        /// <summary>
        /// This element specifies the properties of the fraction object f. Properties of the Fraction object include the type
        /// or style of the fraction.The fraction bar can be horizontal or diagonal, depending on the fraction properties.The
        /// fraction object is also used to represent the stack function, which places one element above another, with no
        /// fraction bar.
        /// </summary>
        public ST_FType FractionType
        {
            get
            {
                return f.fPr.type.val;
            }
            set
            {
                f.fPr.type.val = value;
            }
        }

        //<inheritdoc/>
        public XWPFDocument Document { get { return parent.Document; } }

        //<inheritdoc/>
        public POIXMLDocumentPart Part { get { return parent.Part; } }

        /// <summary>
        /// This element specifies the numerator of the Fraction object f
        /// </summary>
        public XWPFOMathArg Numerator { get { return num; } }

        /// <summary>
        /// This element specifies the denominator of a fraction
        /// </summary>
        public XWPFOMathArg Denominator { get { return den; } }
    }
}
