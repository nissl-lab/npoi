using System;
using System.Collections.Generic;
using System.Globalization;
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
            Write(stream, value);
        }

        internal static void Write<T>(Stream stream, T value)
        {
            using var writer = XmlWriter.Create(stream, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                CloseOutput = false
            });
            new XmlSerializer(typeof(T)).Serialize(writer, value);
        }

        internal static T GetReferencedPart<T>(POIXMLDocumentPart owner, string id, XSLFRelation descriptor)
            where T : POIXMLDocumentPart
        {
            var relationship = string.IsNullOrEmpty(id) ? null : owner.GetPackagePart().GetRelationship(id);
            if (relationship == null || relationship.TargetMode != TargetMode.Internal ||
                relationship.RelationshipType != descriptor.Relation ||
                !(owner.GetRelationById(id) is T part) ||
                part.GetPackagePart().ContentType != descriptor.ContentType ||
                !ReferenceEquals(owner.GetPackagePart().GetRelatedPart(relationship), part.GetPackagePart()))
            {
                throw new POIXMLException("Invalid " + descriptor.Relation + " reference '" + id +
                    "' in " + owner.GetPackagePart().PartName.Name);
            }
            return part;
        }

        internal static List<T> GetReferencedParts<T>(
            POIXMLDocumentPart owner, XmlElement list, string entryName, XSLFRelation descriptor)
            where T : POIXMLDocumentPart
        {
            var parts = new List<T>();
            if (list == null)
            {
                return parts;
            }
            var identifiers = new HashSet<uint>();
            var referencedParts = new HashSet<T>();
            foreach (XmlNode node in list.ChildNodes)
            {
                if (!(node is XmlElement entry) || entry.LocalName != entryName ||
                    entry.NamespaceURI != CT_Presentation.Namespace)
                {
                    continue;
                }
                if (!uint.TryParse(entry.GetAttribute("id"), NumberStyles.Integer, CultureInfo.InvariantCulture,
                    out var identifier) || identifier < 2147483648U || !identifiers.Add(identifier))
                {
                    throw new POIXMLException("Invalid or duplicate " + entryName + " identifier in " +
                        owner.GetPackagePart().PartName.Name);
                }
                var part = GetReferencedPart<T>(owner,
                    entry.GetAttribute("id", CT_Presentation.RelationshipsNamespace), descriptor);
                if (!referencedParts.Add(part))
                {
                    throw new POIXMLException("Duplicate " + entryName + " part reference in " +
                        owner.GetPackagePart().PartName.Name);
                }
                parts.Add(part);
            }
            return parts;
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
