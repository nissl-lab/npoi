using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Open Packaging Convention content types (see Annex F : Standard Namespaces
     * and Content Types).
     *
     * @author CDubettier define some constants, Julien Chable
     */
    public class ContentTypes
    {

        /*
         * Open Packaging Convention (Annex F : Standard Namespaces and Content
         * Types)
         */

        /**
         * Core Properties part.
         */
        public static String CORE_PROPERTIES_PART = "application/vnd.openxmlformats-package.core-properties+xml";

        /**
         * Digital Signature Certificate part.
         */
        public static String DIGITAL_SIGNATURE_CERTIFICATE_PART = "application/vnd.openxmlformats-package.digital-signature-certificate";

        /**
         * Digital Signature Origin part.
         */
        public static String DIGITAL_SIGNATURE_ORIGIN_PART = "application/vnd.openxmlformats-package.digital-signature-origin";

        /**
         * Digital Signature XML Signature part.
         */
        public static String DIGITAL_SIGNATURE_XML_SIGNATURE_PART = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";

        /**
         * Relationships part.
         */
        public static String RELATIONSHIPS_PART = "application/vnd.openxmlformats-package.relationships+xml";

        /**
         * Custom XML part.
         */
        public static String CUSTOM_XML_PART = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

        /**
         * Plain old xml. Note - OOXML uses application/xml, and not text/xml!
         */
        public static String PLAIN_OLD_XML = "application/xml";

        public static String IMAGE_JPEG = "image/jpeg";

        public static String EXTENSION_JPG_1 = "jpg";

        public static String EXTENSION_JPG_2 = "jpeg";

        // image/png ISO/IEC 15948:2003 http://www.libpng.org/pub/png/spec/
        public static String IMAGE_PNG = "image/png";

        public static String EXTENSION_PNG = "png";

        // image/gif http://www.w3.org/Graphics/GIF/spec-gif89a.txt
        public static String IMAGE_GIF = "image/gif";

        public static String EXTENSION_GIF = "gif";

        /**
         * TIFF image format.
         *
         * @see <a href="http://partners.adobe.com/public/developer/tiff/index.html#spec">
         * http://partners.adobe.com/public/developer/tiff/index.html#spec</a>
         */
        public static String IMAGE_TIFF = "image/tiff";

        public static String EXTENSION_TIFF = "tiff";

        /**
         * Pict image format.
         *
         * @see <a href="http://developer.apple.com/documentation/mac/QuickDraw/QuickDraw-2.html">
         * http://developer.apple.com/documentation/mac/QuickDraw/QuickDraw-2.html</a>
         */
        public static String IMAGE_PICT = "image/pict";

        public static String EXTENSION_PICT = "tiff";

        /**
         * XML file.
         */
        public static String XML = "text/xml";

        public static String EXTENSION_XML = "xml";

        public static String GetContentTypeFromFileExtension(String filename)
        {
            String extension = filename.Substring(filename.LastIndexOf(".") + 1)
                    .ToLower();
            if (extension.Equals(EXTENSION_JPG_1)
                    || extension.Equals(EXTENSION_JPG_2))
                return IMAGE_JPEG;
            else if (extension.Equals(EXTENSION_GIF))
                return IMAGE_GIF;
            else if (extension.Equals(EXTENSION_PICT))
                return IMAGE_PICT;
            else if (extension.Equals(EXTENSION_PNG))
                return IMAGE_PNG;
            else if (extension.Equals(EXTENSION_TIFF))
                return IMAGE_TIFF;
            else if (extension.Equals(EXTENSION_XML))
                return XML;
            else
                return null;
        }
    }
}
