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
    /// Relationship types.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 0.2
    /// </remarks>

    public static class PackageRelationshipTypes
    {    
        /// <summary>
        /// <para>
        /// Core properties relationship type.
        /// </para>
        /// <para>
        ///  The standard specifies a source relations ship for the Core File Properties part as follows:
        ///  <c>http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties.</c>
        /// </para>
        /// <para>
        /// </para>
        /// <para>
        ///   Office uses the following source relationship for the Core File Properties part:
        ///   <c>http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties.</c>
        /// </para>
        /// <para>
        /// See 2.1.33 Part 1 Section 15.2.11.1, Core File Properties Part in [MS-OE376].pdf
        /// </para>
        /// </summary>
        public const string CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

        /// <summary>
        /// Core properties relationship type as defiend in ECMA 376.
        /// </summary>
        public const string CORE_PROPERTIES_ECMA376 = "http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties";

        /// <summary>
        /// Namespace of Core properties relationship type as defiend in ECMA 376
        /// </summary>
        public const string CORE_PROPERTIES_ECMA376_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>
        /// Digital signature relationship type.
        /// </summary>
        public const string DIGITAL_SIGNATURE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

        /// <summary>
        /// Digital signature certificate relationship type.
        /// </summary>
        public const string DIGITAL_SIGNATURE_CERTIFICATE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";

        /// <summary>
        /// Digital signature origin relationship type.
        /// </summary>
        public const string DIGITAL_SIGNATURE_ORIGIN = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";

        /// <summary>
        /// Thumbnail relationship type.
        /// </summary>
        public const string THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

        /// <summary>
        /// Extended properties relationship type.
        /// </summary>
        public const string EXTENDED_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

        /// <summary>
        /// Extended properties relationship type for strict ooxml.
        /// </summary>
        public const string STRICT_EXTENDED_PROPERTIES = "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties";

        /// <summary>
        /// Custom properties relationship type.
        /// </summary>
        public const string CUSTOM_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

        /// <summary>
        /// Core document relationship type.
        /// </summary>
        public const string CORE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

        /// <summary>
        /// Core document relationship type for strict ooxml.
        /// </summary>
        public const string STRICT_CORE_DOCUMENT = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";

        /// <summary>
        /// Custom XML relationship type.
        /// </summary>
        public const string CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";

        /// <summary>
        /// Image type.
        /// </summary>
        public const string IMAGE_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        /// <summary>
        /// Hyperlink type.
        /// </summary>
        public const string HYPERLINK_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        /// <summary>
        /// Style type.
        /// </summary>
        public const string STYLE_PART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

        /// <summary>
        /// External Link to another Document
        /// </summary>
        public const string EXTERNAL_LINK_PATH = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

        /// <summary>
        /// Visio 2010 VSDX equivalent of package <see cref="CORE_DOCUMENT"/>
        /// </summary>
        public const string VISIO_CORE_DOCUMENT = "http://schemas.microsoft.com/visio/2010/relationships/document";
    }

}
