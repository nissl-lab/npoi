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
    {        /**
         * Core properties relationship type.
         */
        public const String CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

        /**
         * Digital signature relationship type.
         */
        public const String DIGITAL_SIGNATURE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

        /**
         * Digital signature certificate relationship type.
         */
        public const String DIGITAL_SIGNATURE_CERTIFICATE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";

        /**
         * Digital signature origin relationship type.
         */
        public const String DIGITAL_SIGNATURE_ORIGIN = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";

        /**
         * Thumbnail relationship type.
         */
        public const String THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

        /**
         * Extended properties relationship type.
         */
        public const String EXTENDED_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

        /**
         * Custom properties relationship type.
         */
        public const String CUSTOM_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
	
        /**
         * Core properties relationship type.
         */
        public const String CORE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

        /**
         * Custom XML relationship type.
         */
        public const String CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";

        /**
         * Image type.
         */
        public const String IMAGE_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        /**
         * Style type.
         */
        public const String STYLE_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    }

}
