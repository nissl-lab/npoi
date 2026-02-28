using NPOI.OpenXmlFormats.Shared;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.XWPF.Usermodel
{
    /// <summary>
    /// n-ary Operator Object
    /// This element specifies an n-ary object, consisting of an n-ary object, a base (or operand), and optional upper and
    /// lower limits
    /// </summary>
    public class XWPFNary:IRunBody
    {
        private CT_Nary nary;
        private IRunBody parent;
        private readonly XWPFOMathArg e;
        private readonly XWPFOMathArg sub;
        private readonly XWPFOMathArg sup;

        public XWPFNary(CT_Nary nary, IRunBody p)
        {
            this.nary = nary;
            this.parent = p;

            if (nary.e == null)
            {
                nary.e = new CT_OMathArg();
            }
            this.e = new XWPFOMathArg(nary.e, this);

            if (nary.sub == null)
            {
                nary.sub = new CT_OMathArg();
            }
            this.sub = new XWPFOMathArg(nary.sub, this);

            if (nary.sup == null)
            {
                nary.sup = new CT_OMathArg();
            }
            this.sup = new XWPFOMathArg(nary.sup, this);

            if(nary.naryPr == null)
            {
                nary.naryPr = new CT_NaryPr();
            }

            if (nary.naryPr.chr == null)
            {
                nary.naryPr.chr = new CT_Char();
            }
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

        /// <summary>
        /// This element specifies the superscript of the superscript object sSup. 
        /// </summary>
        public XWPFOMathArg Superscript { get { return sup; } }

        /// <summary>
        /// Get Nary symbol
        /// </summary>
        public string NaryPr
        {
            get { return nary.naryPr.chr.val; }
        }

        /// <summary>
        /// Sets ∑ char
        /// </summary>
        public XWPFNary SetSumm()
        {
            nary.naryPr.chr.val = "∑";
            return this;
        }

        /// <summary>
        /// Sets ⋃ char
        /// </summary>
        public XWPFNary SetUnion()
        {
            nary.naryPr.chr.val = "⋃";
            return this;
        }

        /// <summary>
        /// Sets ∫ char
        /// </summary>
        public XWPFNary SetIntegral()
        {
            nary.naryPr.chr.val = "∫";
            return this;
        }

        /// <summary>
        /// Sets ⋀ char
        /// </summary>
        public XWPFNary SetAnd()
        {
            nary.naryPr.chr.val = "⋀";
            return this;
        }

    }
}
