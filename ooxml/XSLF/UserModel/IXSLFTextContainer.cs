using System.Xml;

namespace NPOI.XSLF.UserModel
{
    internal interface IXSLFTextContainer
    {
        XmlElement GetXmlObject();
        XSLFSheet GetSheet();
        XmlElement GetTextBody(bool create);
        void EnsureTextEditable();
    }
}
