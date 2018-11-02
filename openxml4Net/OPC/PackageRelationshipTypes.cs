using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Relationship types.
     *
     * @author Julien Chable
     * @version 0.2
     */
    public static class PackageRelationshipTypes
    {    
        /**
         * Core properties relationship type.
         *
         *  <p>
         *  The standard specifies a source relations ship for the Core File Properties part as follows:
         *  <code>http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties.</code>
         *  </p>
         *  <p>
         *   Office uses the following source relationship for the Core File Properties part:
         *   <code>http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties.</code>
         * </p>
         * See 2.1.33 Part 1 Section 15.2.11.1, Core File Properties Part in [MS-OE376].pdf
         */
        public const string CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

        /**
         * Core properties relationship type as defiend in ECMA 376.
         */
        public const string CORE_PROPERTIES_ECMA376 = "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties";

        /**
         * Digital signature relationship type.
         */
        public const string DIGITAL_SIGNATURE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

        /**
         * Digital signature certificate relationship type.
         */
        public const string DIGITAL_SIGNATURE_CERTIFICATE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";

        /**
         * Digital signature origin relationship type.
         */
        public const string DIGITAL_SIGNATURE_ORIGIN = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";

        /**
         * Thumbnail relationship type.
         */
        public const string THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

        /**
         * Extended properties relationship type.
         */
        public const string EXTENDED_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

        /**
         * Extended properties relationship type for strict ooxml.
         */
        public const string STRICT_EXTENDED_PROPERTIES = "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties";

        /**
         * Custom properties relationship type.
         */
        public const string CUSTOM_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

        /**
         * Core document relationship type.
         */
        public const string CORE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

        /**
         * Core document relationship type for strict ooxml.
         */
        public const string STRICT_CORE_DOCUMENT = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";

        /**
         * Custom XML relationship type.
         */
        public const string CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";

        /**
         * Image type.
         */
        public const string IMAGE_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        /**
         * Hyperlink type.
         */
        public const string HYPERLINK_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        /**
         * Style type.
         */
        public const string STYLE_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

        /**
         * External Link to another Document
         */
        public const string EXTERNAL_LINK_PATH = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

        /**
         * Visio 2010 VSDX equivalent of package {@link #CORE_DOCUMENT}
         */
        public const string VISIO_CORE_DOCUMENT = "http://schemas.microsoft.com/visio/2010/relationships/document";
    }

}
