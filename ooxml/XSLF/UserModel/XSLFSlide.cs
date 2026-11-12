using System.IO;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A PresentationML slide with shape containers and preserved unsupported drawing content.
    /// </summary>
    public class XSLFSlide : XSLFSheet
    {
        private readonly CT_Slide slide;

        internal XSLFSlide()
        {
            slide = new CT_Slide
            {
                Namespaces = XSLFXmlHelper.CreateNamespaces(),
                cSld = XSLFXmlHelper.CreateCommonSlideData(),
                AnyElements = new[]
                {
                    XSLFXmlHelper.CreateElement(
                        "<p:clrMapOvr xmlns:p=\"" + CT_Presentation.Namespace + "\" xmlns:a=\"" +
                        XSLFXmlHelper.DrawingNamespace + "\"><a:masterClrMapping/></p:clrMapOvr>")
                }
            };
        }

        internal XSLFSlide(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {
            slide = XSLFXmlHelper.Read<CT_Slide>(part);
            if (slide.cSld == null)
            {
                throw new POIXMLException("Slide has no common slide data: " + part.PartName.Name);
            }
            InitializeXml();
        }

        /// <summary>
        /// Returns the underlying partial PresentationML model.
        /// </summary>
        public CT_Slide GetXmlObject()
        {
            return slide;
        }

        protected override XmlElement GetCommonSlideData()
        {
            return slide.cSld;
        }

        /// <summary>
        /// Returns the slide layout referenced by this slide.
        /// </summary>
        public XSLFSlideLayout GetSlideLayout()
        {
            return GetRequiredRelation<XSLFSlideLayout>(XSLFRelation.SLIDE_LAYOUT);
        }

        /// <summary>
        /// Returns the slide master inherited through this slide's layout.
        /// </summary>
        public XSLFSlideMaster GetSlideMaster()
        {
            return GetSlideLayout().GetSlideMaster();
        }

        /// <inheritdoc/>
        public override XSLFSheet GetMasterSheet()
        {
            return GetSlideLayout();
        }

        /// <inheritdoc/>
        public override XSLFTheme GetTheme()
        {
            return GetSlideLayout().GetTheme();
        }

        protected override void WriteXml(Stream stream)
        {
            XSLFXmlHelper.Write(stream, slide);
        }
    }
}
