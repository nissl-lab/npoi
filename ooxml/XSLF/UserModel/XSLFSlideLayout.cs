using System.IO;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A slide layout and its relationship to a slide master.
    /// </summary>
    public class XSLFSlideLayout : XSLFSheet
    {
        private readonly CT_SlideLayout layout;

        internal XSLFSlideLayout(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {
            layout = XSLFXmlHelper.Read<CT_SlideLayout>(part);
            if (layout.cSld == null)
            {
                throw new POIXMLException("Slide layout has no common slide data: " + part.PartName.Name);
            }
            InitializeXml();
        }

        /// <summary>
        /// Gets the layout name, or an empty string if it has no name.
        /// </summary>
        public string Name
        {
            get { return layout.cSld.GetAttribute("name"); }
        }

        /// <summary>
        /// Gets the PresentationML layout type, such as "blank" or "title".
        /// </summary>
        public string Type
        {
            get { return layout.type ?? "cust"; }
        }

        /// <summary>
        /// Returns the underlying partial PresentationML model.
        /// </summary>
        public CT_SlideLayout GetXmlObject()
        {
            return layout;
        }

        protected override XmlElement GetCommonSlideData()
        {
            return layout.cSld;
        }

        /// <summary>
        /// Returns the slide master referenced by this layout.
        /// </summary>
        public XSLFSlideMaster GetSlideMaster()
        {
            return GetRequiredRelation<XSLFSlideMaster>(XSLFRelation.SLIDE_MASTER);
        }

        /// <inheritdoc/>
        public override XSLFSheet GetMasterSheet()
        {
            return GetSlideMaster();
        }

        /// <inheritdoc/>
        public override XSLFTheme GetTheme()
        {
            return GetSlideMaster().GetTheme();
        }

        protected override void WriteXml(Stream stream)
        {
            XSLFXmlHelper.Write(stream, layout);
        }
    }
}
