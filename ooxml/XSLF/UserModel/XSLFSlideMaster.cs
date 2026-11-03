using System;
using System.IO;
using System.Linq;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Presentation;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// A presentation slide master and its referenced layouts.
    /// </summary>
    public class XSLFSlideMaster : XSLFSheet
    {
        private readonly CT_SlideMaster master;

        internal XSLFSlideMaster(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {
            master = XSLFXmlHelper.Read<CT_SlideMaster>(part);
            if (master.cSld == null || master.clrMap == null)
            {
                throw new POIXMLException("Slide master is missing common slide data or its color map: " +
                    part.PartName.Name);
            }
            InitializeXml();
        }

        /// <summary>
        /// Gets the name of the common slide data, or an empty string if it has no name.
        /// </summary>
        public string Name
        {
            get { return master.cSld.GetAttribute("name"); }
        }

        /// <summary>
        /// Returns the underlying partial PresentationML model.
        /// </summary>
        public CT_SlideMaster GetXmlObject()
        {
            return master;
        }

        /// <summary>
        /// Returns layouts in the order of the master's layout identifier list.
        /// </summary>
        public XSLFSlideLayout[] GetSlideLayouts()
        {
            return XSLFXmlHelper.GetReferencedParts<XSLFSlideLayout>(
                this, master.sldLayoutIdLst, "sldLayoutId", XSLFRelation.SLIDE_LAYOUT).ToArray();
        }

        /// <summary>
        /// Finds a layout by name, ignoring case, or returns null if no layout matches.
        /// </summary>
        public XSLFSlideLayout GetLayout(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }
            return GetSlideLayouts().FirstOrDefault(layout =>
                string.Equals(layout.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <inheritdoc/>
        public override XSLFSheet GetMasterSheet()
        {
            return null;
        }

        /// <inheritdoc/>
        public override XSLFTheme GetTheme()
        {
            if (GetPackagePart().GetRelationshipsByType(XSLFRelation.THEME.Relation).Size == 0)
            {
                return null;
            }
            return GetRequiredRelation<XSLFTheme>(XSLFRelation.THEME);
        }

        protected override void WriteXml(Stream stream)
        {
            XSLFXmlHelper.Write(stream, master);
        }
    }
}
