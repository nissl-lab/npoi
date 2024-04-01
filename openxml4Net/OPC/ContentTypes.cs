/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// Open Packaging Convention content types (see Annex F : Standard Namespaces
    /// and Content Types).
    /// </summary>
    /// <remarks>
    /// @author CDubettier define some constants, Julien Chable
    /// </remarks>

    public class ContentTypes
    {

        /*
         * Open Packaging Convention (Annex F : Standard Namespaces and Content
         * Types)
         */

        /// <summary>
        /// Core Properties part.
        /// </summary>
        public static String CORE_PROPERTIES_PART = "application/vnd.openxmlformats-package.core-properties+xml";

        /// <summary>
        /// Digital Signature Certificate part.
        /// </summary>
        public static String DIGITAL_SIGNATURE_CERTIFICATE_PART = "application/vnd.openxmlformats-package.digital-signature-certificate";

        /// <summary>
        /// Digital Signature Origin part.
        /// </summary>
        public static String DIGITAL_SIGNATURE_ORIGIN_PART = "application/vnd.openxmlformats-package.digital-signature-origin";

        /// <summary>
        /// Digital Signature XML Signature part.
        /// </summary>
        public static String DIGITAL_SIGNATURE_XML_SIGNATURE_PART = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";

        /// <summary>
        /// Relationships part.
        /// </summary>
        public static String RELATIONSHIPS_PART = "application/vnd.openxmlformats-package.relationships+xml";

        /// <summary>
        /// Custom XML part.
        /// </summary>
        public static String CUSTOM_XML_PART = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

        /// <summary>
        /// Plain old xml. Note - OOXML uses application/xml, and not text/xml!
        /// </summary>
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

        /// <summary>
        /// <para>
        /// TIFF image format.
        /// </para>
        /// <para>
        /// <see href="http://partners.adobe.com/public/developer/tiff/index.html#spec"> http://partners.adobe.com/public/developer/tiff/index.html#spec</see>
        /// </para>
        /// </summary>
        public static String IMAGE_TIFF = "image/tiff";

        public static String EXTENSION_TIFF = "tiff";

        /// <summary>
        /// <para>
        /// Pict image format.
        /// </para>
        /// <para>
        /// <see href="http://developer.apple.com/documentation/mac/QuickDraw/QuickDraw-2.html"> http://developer.apple.com/documentation/mac/QuickDraw/QuickDraw-2.html</see>
        /// </para>
        /// </summary>
        public static String IMAGE_PICT = "image/pict";

        public static String EXTENSION_PICT = "tiff";

        /// <summary>
        /// XML file.
        /// </summary>
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
