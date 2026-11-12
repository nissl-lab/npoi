using System;
using System.Xml;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Preserves an unsupported drawing element without interpreting its geometry or relationships.
    /// Common nonvisual metadata can be read when unambiguous; editing is not supported.
    /// </summary>
    public class XSLFUnsupportedShape : XSLFShape
    {
        internal XSLFUnsupportedShape(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
            : base(xml, sheet, parent)
        {
        }

        internal override void EnsureEditable()
        {
            throw new NotSupportedException("Opaque drawing elements do not support editing.");
        }
    }
}
