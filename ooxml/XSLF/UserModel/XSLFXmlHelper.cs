using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.Util;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    internal static class XSLFXmlHelper
    {
        internal const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

        internal static T Read<T>(PackagePart part)
        {
            try
            {
                using var stream = part.GetInputStream();
                using var reader = XmlReader.Create(stream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null
                });
                return (T)new XmlSerializer(typeof(T)).Deserialize(reader);
            }
            catch (InvalidOperationException e)
            {
                throw new POIXMLException("Invalid PresentationML in " + part.PartName.Name, e);
            }
        }

        internal static void Write<T>(PackagePart part, T value)
        {
            using var stream = part.GetOutputStream();
            using var writer = XmlWriter.Create(stream, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                CloseOutput = false
            });
            new XmlSerializer(typeof(T)).Serialize(writer, value);
        }

        internal static XmlSerializerNamespaces CreateNamespaces()
        {
            var namespaces = new XmlSerializerNamespaces();
            namespaces.Add("p", CT_Presentation.Namespace);
            namespaces.Add("a", DrawingNamespace);
            namespaces.Add("r", CT_Presentation.RelationshipsNamespace);
            return namespaces;
        }

        internal static XmlElement CreateElement(string xml)
        {
            var document = new XmlDocument { XmlResolver = null };
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            XmlHelper.LoadXmlSafe(document, stream);
            return document.DocumentElement;
        }

        internal static XmlElement CreateCommonSlideData()
        {
            return CreateElement(
                "<p:cSld xmlns:p=\"" + CT_Presentation.Namespace + "\">" +
                "<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/>" +
                "<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>");
        }
    }
}
