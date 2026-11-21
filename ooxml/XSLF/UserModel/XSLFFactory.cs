using System;
using NPOI.OpenXml4Net.OPC;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Creates slide, master, layout, and theme wrappers while preserving unmodeled parts.
    /// </summary>
    public class XSLFFactory : POIXMLFactory
    {
        private static readonly XSLFFactory Instance = new XSLFFactory();

        private XSLFFactory()
        {
        }

        /// <summary>
        /// Returns the shared XSLF factory.
        /// </summary>
        public static XSLFFactory GetInstance()
        {
            return Instance;
        }

        /// <inheritdoc/>
        public override POIXMLDocumentPart CreateDocumentPart(POIXMLDocumentPart parent, PackagePart part)
        {
            if (part.ContentType == XSLFRelation.SLIDE.ContentType)
            {
                return new XSLFSlide(parent, part);
            }
            if (part.ContentType == XSLFRelation.SLIDE_MASTER.ContentType)
            {
                return new XSLFSlideMaster(parent, part);
            }
            if (part.ContentType == XSLFRelation.SLIDE_LAYOUT.ContentType)
            {
                return new XSLFSlideLayout(parent, part);
            }
            if (part.ContentType == XSLFRelation.THEME.ContentType)
            {
                return new XSLFTheme(parent, part);
            }

            // The default POIXML save traversal clears parts before committing them.
            // Unmodeled parts have no serializer and must keep their original bytes.
            return new POIXMLDocumentPart(parent, part) { Committed = true };
        }

        protected override POIXMLRelation GetDescriptor(string relationshipType)
        {
            return XSLFRelation.GetInstance(relationshipType);
        }

        protected override POIXMLDocumentPart CreateDocumentPart(Type cls, Type[] classes, object[] values)
        {
            if (cls == typeof(XSLFSlide) && classes == null)
            {
                return new XSLFSlide();
            }
            throw new NotSupportedException("The XSLF factory cannot create this document part.");
        }
    }
}
