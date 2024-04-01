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
    /// Open Packaging Convention namespaces URI.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public class PackageNamespaces
    {
        /// <summary>
        /// Dublin Core Terms URI.
        /// </summary>
        public const String NAMESPACE_DCTERMS = "http://purl.org/dc/terms/";

        /// <summary>
        /// Dublin Core namespace URI.
        /// </summary>
        public const String NAMESPACE_DC = "http://purl.org/dc/elements/1.1/";
        /// <summary>
        /// Content Types.
        /// </summary>
        public const String CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

        /// <summary>
        /// Core Properties.
        /// </summary>
        public const String CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

        /// <summary>
        /// Digital Signatures.
        /// </summary>
        public const String DIGITAL_SIGNATURE = "http://schemas.openxmlformats.org/package/2006/digital-signature";

        /// <summary>
        /// Relationships.
        /// </summary>
        public const String RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>
        /// Markup Compatibility.
        /// </summary>
        public const String MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public const string DCMITYPE = "http://purl.org/dc/dcmitype/";

        public const string SCHEMA_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        public const string SCHEMA_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public const string SCHEMA_SHEETDRAWINGS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        public const string SCHEMA_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        public const string SCHEMA_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        public const string SCHEMA_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    }

}
