using System.Xml;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A simple shape explicitly marked as a text box.
    /// </summary>
    public class XSLFTextBox : XSLFSimpleShape
    {
        internal XSLFTextBox(XmlElement xml, XSLFSheet sheet, IXSLFShapeContainer parent)
            : base(xml, sheet, parent)
        {
        }
    }
}
