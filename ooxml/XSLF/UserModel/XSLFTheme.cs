using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A read-only view of theme metadata backed by the existing DrawingML theme model.
    /// </summary>
    public class XSLFTheme : POIXMLDocumentPart
    {
        private readonly CT_OfficeStyleSheet theme;

        internal XSLFTheme(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {
            var document = new XmlDocument { XmlResolver = null };
            using var stream = part.GetInputStream();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            try
            {
                document.Load(reader);
                var root = document.DocumentElement;
                if (root == null || root.LocalName != "theme" ||
                    root.NamespaceURI != XSLFXmlHelper.DrawingNamespace)
                {
                    throw new POIXMLException("Invalid DrawingML theme: " + part.PartName.Name);
                }
                theme = CT_OfficeStyleSheet.Parse(root, NamespaceManager);
            }
            catch (XmlException e)
            {
                throw new POIXMLException("Invalid DrawingML theme: " + part.PartName.Name, e);
            }
            Committed = true;
        }

        /// <summary>
        /// Gets the theme name.
        /// </summary>
        public string Name
        {
            get { return theme.name; }
        }

        /// <summary>
        /// Gets the major Latin font, or null if it is not specified.
        /// </summary>
        public string MajorFont
        {
            get { return theme.themeElements?.fontScheme?.majorFont?.latin?.typeface; }
        }

        /// <summary>
        /// Gets the minor Latin font, or null if it is not specified.
        /// </summary>
        public string MinorFont
        {
            get { return theme.themeElements?.fontScheme?.minorFont?.latin?.typeface; }
        }
    }
}
