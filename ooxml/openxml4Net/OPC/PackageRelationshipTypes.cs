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
    public class PackageRelationshipTypes
    {
        /**
         * Core properties relationship type.
         */
        public static String CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

        /**
         * Digital signature relationship type.
         */
        public static String DIGITAL_SIGNATURE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

        /**
         * Digital signature certificate relationship type.
         */
        public static String DIGITAL_SIGNATURE_CERTIFICATE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";

        /**
         * Digital signature origin relationship type.
         */
        public static String DIGITAL_SIGNATURE_ORIGIN = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";

        /**
         * Thumbnail relationship type.
         */
        public static String THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

        /**
         * Extended properties relationship type.
         */
        public static String EXTENDED_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

        /**
         * Core properties relationship type.
         */
        public static String CORE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

        /**
         * Custom XML relationship type.
         */
        public static String CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";

        /**
         * Image type.
         */
        public static String IMAGE_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        /**
         * Style type.
         */
        public static String STYLE_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    }

}
