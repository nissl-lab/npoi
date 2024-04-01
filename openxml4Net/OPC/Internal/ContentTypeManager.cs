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
using NPOI.OpenXml4Net.Exceptions;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /// <summary>
    /// Manage package content types ([Content_Types].xml part).
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public abstract class ContentTypeManager
    {

        /// <summary>
        /// Content type part name.
        /// </summary>
        public const String CONTENT_TYPES_PART_NAME = "[Content_Types].xml";

        /// <summary>
        /// Content type namespace
        /// </summary>
        public const String TYPES_NAMESPACE_URI = "http://schemas.openxmlformats.org/package/2006/content-types";

        /* Xml elements in content type part */

        private const String TYPES_TAG_NAME = "Types";

        private const String DEFAULT_TAG_NAME = "Default";

        private const String EXTENSION_ATTRIBUTE_NAME = "Extension";

        private const String CONTENT_TYPE_ATTRIBUTE_NAME = "ContentType";

        private const String OVERRIDE_TAG_NAME = "Override";

        private const String PART_NAME_ATTRIBUTE_NAME = "PartName";

        /// <summary>
        /// Reference to the package using this content type manager.
        /// </summary>
        protected OPCPackage container;

        /// <summary>
        /// Default content type tree. <Extension, ContentType>
        /// </summary>
        private SortedList<String, String> defaultContentType;

        /// <summary>
        /// Override content type tree.
        /// </summary>
        private SortedList<PackagePartName, String> overrideContentType;

        /// <summary>
        /// Constructor. Parses the content of the specified input stream.
        /// </summary>
        /// <param name="in">If different of <i>null</i> then the content types part is
        /// retrieve and parse.
        /// </param>
        /// <exception cref="InvalidFormatException">If the content types part content is not valid.
        /// </exception>
        public ContentTypeManager(Stream in1, OPCPackage pkg)
        {
            this.container = pkg;
            this.defaultContentType = new SortedList<String, String>();
            if (in1 != null)
            {
                try
                {
                    ParseContentTypesFile(in1);
                }
                catch (InvalidFormatException ex)
                {
                    throw new InvalidFormatException("Can't read content types part !", ex);
                }
            }
        }

        /// <summary>
        /// <para>
        /// Build association extention-> content type (will be stored in
        /// [Content_Types].xml) for example ContentType="image/png" Extension="png"
        /// </para>
        /// <para>
        /// [M2.8]: When adding a new part to a package, the package implementer
        /// shall ensure that a content type for that part is specified in the
        /// Content Types stream; the package implementer shall perform the steps
        /// described in &#167;9.1.2.3:
        /// </para>
        /// <para>
        /// 1. Get the extension from the part name by taking the substring to the
        /// right of the rightmost occurrence of the dot character (.) from the
        /// rightmost segment.
        /// </para>
        /// <para>
        /// 2. If a part name has no extension, a corresponding Override element
        /// shall be added to the Content Types stream.
        /// </para>
        /// <para>
        /// 3. Compare the resulting extension with the values specified for the
        /// Extension attributes of the Default elements in the Content Types stream.
        /// The comparison shall be case-insensitive ASCII.
        /// </para>
        /// <para>
        /// 4. If there is a Default element with a matching Extension attribute,
        /// then the content type of the new part shall be compared with the value of
        /// the ContentType attribute. The comparison might be case-sensitive and
        /// include every character regardless of the role it plays in the
        /// content-type grammar of RFC 2616, or it might follow the grammar of RFC
        /// 2616.
        /// </para>
        /// <para>
        /// a. If the content types match, no further action is required.
        /// </para>
        /// <para>
        /// b. If the content types do not match, a new Override element shall be
        /// added to the Content Types stream. .
        /// </para>
        /// <para>
        /// 5. If there is no Default element with a matching Extension attribute, a
        /// new Default element or Override element shall be added to the Content
        /// Types stream.
        /// </para>
        /// </summary>
        public void AddContentType(PackagePartName partName, String contentType)
        {
            bool defaultCTExists = false;
            String extension = partName.Extension.ToLower();
            if ((extension.Length == 0)
                    || (this.defaultContentType.ContainsKey(extension) && !(defaultCTExists = this.defaultContentType
                            .ContainsValue(contentType))))
                this.AddOverrideContentType(partName, contentType);
            else if (!defaultCTExists)
                this.AddDefaultContentType(extension, contentType);
        }

        /// <summary>
        /// Add an override content type for a specific part.
        /// </summary>
        /// <param name="partName">
        /// Name of the part.
        /// </param>
        /// <param name="contentType">
        /// Content type of the part.
        /// </param>
        private void AddOverrideContentType(PackagePartName partName,
                String contentType)
        {
            if (overrideContentType == null)
                overrideContentType = new SortedList<PackagePartName, String>();

            if(!overrideContentType.ContainsKey(partName))
                overrideContentType.Add(partName, contentType);
            else
                overrideContentType[partName]= contentType;
        }

        /// <summary>
        /// Add a content type associated with the specified extension.
        /// </summary>
        /// <param name="extension">
        /// The part name extension to bind to a content type.
        /// </param>
        /// <param name="contentType">
        /// The content type associated with the specified extension.
        /// </param>
        private void AddDefaultContentType(String extension, String contentType)
        {
            // Remark : Originally the latest parameter was :
            // contentType.toLowerCase(). Change due to a request ID 1996748.
            defaultContentType.Add(extension.ToLower(), contentType);
        }

        /// <summary>
        /// <para>
        /// Delete a content type based on the specified part name. If the specified
        /// part name is register with an override content type, then this content
        /// type is remove, else the content type is remove in the default content
        /// type list if it exists and if no part is associated with it yet.
        /// </para>
        /// <para>
        /// Check rule M2.4: The package implementer shall require that the Content
        /// Types stream contain one of the following for every part in the package:
        /// One matching Default element One matching Override element Both a
        /// matching Default element and a matching Override element, in which case
        /// the Override element takes precedence.
        /// </para>
        /// </summary>
        /// <param name="partName">
        /// The part URI associated with the override content type to
        /// delete.
        /// </param>
        /// <exception cref="InvalidOperationException">InvalidOperationException
        /// Throws if
        /// </exception>
        public void RemoveContentType(PackagePartName partName)
        {
            if (partName == null)
                throw new ArgumentException("partName");

            /* Override content type */
            if (this.overrideContentType != null
                    && this.overrideContentType.ContainsKey(partName))
            {
                // Remove the override definition for the specified part.
                this.overrideContentType.Remove(partName);
                return;
            }

            /* Default content type */
            String extensionToDelete = partName.Extension;
            bool deleteDefaultContentTypeFlag = true;
            if (this.container != null)
            {
                try
                {
                    foreach (PackagePart part in this.container.GetParts())
                    {
                        if (!part.PartName.Equals(partName) && part.PartName.Extension
                                        .Equals(extensionToDelete, StringComparison.InvariantCultureIgnoreCase))
                        {
                            deleteDefaultContentTypeFlag = false;
                            break;
                        }
                    }
                }
                catch (InvalidFormatException e)
                {
                    throw new InvalidOperationException(e.Message);
                }
            }

            // Remove the default content type, no other part use this content type.
            if (deleteDefaultContentTypeFlag)
            {
                this.defaultContentType.Remove(extensionToDelete);
            }

            /*
             * Check rule 2.4: The package implementer shall require that the
             * Content Types stream contain one of the following for every part in
             * the package: One matching Default element One matching Override
             * element Both a matching Default element and a matching Override
             * element, in which case the Override element takes precedence.
             */
            if (this.container != null)
            {
                try
                {
                    foreach (PackagePart part in this.container.GetParts())
                    {
                        if (!part.PartName.Equals(partName)
                                && this.GetContentType(part.PartName) == null)
                            throw new InvalidOperationException(
                                    "Rule M2.4 is not respected: Nor a default element or override element is associated with the part: "
                                            + part.PartName.Name);
                    }
                }
                catch (InvalidFormatException e)
                {
                    throw new InvalidOperationException(e.Message);
                }
            }
        }

        /// <summary>
        /// Check if the specified content type is already register.
        /// </summary>
        /// <param name="contentType">
        /// The content type to check.
        /// </param>
        /// <returns><c>true</c> if the specified content type is already
        /// register, then <c>false</c>.
        /// </returns>
        public bool IsContentTypeRegister(String contentType)
        {
            if (contentType == null)
                throw new ArgumentException("contentType");

            return (this.defaultContentType.Values.Contains(contentType) || (this.overrideContentType != null && this.overrideContentType
                    .Values.Contains(contentType)));
        }

        /// <summary>
        /// <para>
        /// Get the content type for the specified part, if any.
        /// </para>
        /// <para>
        /// Rule [M2.9]: To get the content type of a part, the package implementer
        /// shall perform the steps described in &#167;9.1.2.4:
        /// </para>
        /// <para>
        /// 1. Compare the part name with the values specified for the PartName
        /// attribute of the Override elements. The comparison shall be
        /// case-insensitive ASCII.
        /// </para>
        /// <para>
        /// 2. If there is an Override element with a matching PartName attribute,
        /// return the value of its ContentType attribute. No further action is
        /// required.
        /// </para>
        /// <para>
        /// 3. If there is no Override element with a matching PartName attribute,
        /// then a. Get the extension from the part name by taking the substring to
        /// the right of the rightmost occurrence of the dot character (.) from the
        /// rightmost segment. b. Check the Default elements of the Content Types
        /// stream, comparing the extension with the value of the Extension
        /// attribute. The comparison shall be case-insensitive ASCII.
        /// </para>
        /// <para>
        /// 4. If there is a Default element with a matching Extension attribute,
        /// return the value of its ContentType attribute. No further action is
        /// required.
        /// </para>
        /// <para>
        /// 5. If neither Override nor Default elements with matching attributes are
        /// found for the specified part name, the implementation shall not map this
        /// part name to a part.
        /// </para>
        /// </summary>
        /// <param name="partName">
        /// The URI part to check.
        /// </param>
        /// <returns>The content type associated with the URI (in case of an override
        /// content type) or the extension (in case of default content type),
        /// else <c>null</c>.
        /// </returns>
        /// 
        /// <exception cref="OpenXml4NetException">
        /// Throws if the content type manager is not able to find the
        /// content from an existing part.
        /// </exception>
        public String GetContentType(PackagePartName partName)
        {
            if (partName == null)
                throw new ArgumentException("partName");

            if ((this.overrideContentType != null)
                    && this.overrideContentType.ContainsKey(partName))
                return this.overrideContentType[partName];

            String extension = partName.Extension.ToLower();
            if (this.defaultContentType.ContainsKey(extension))
                return this.defaultContentType[extension];

            /*
             * [M2.4] : The package implementer shall require that the Content Types
             * stream contain one of the following for every part in the package:
             * One matching Default element, One matching Override element, Both a
             * matching Default element and a matching Override element, in which
             * case the Override element takes precedence.
             */
            if (this.container != null && this.container.GetPart(partName) != null)
            {
                throw new OpenXml4NetException(
                        "Rule M2.4 exception : this error should NEVER happen! If you can provide the triggering file, then please raise a bug at https://bz.apache.org/bugzilla/enter_bug.cgi?product=POI and attach a file that triggers it, thanks!");
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Clear all content types.
        /// </summary>
        public void ClearAll()
        {
            this.defaultContentType.Clear();
            if (this.overrideContentType != null)
                this.overrideContentType.Clear();
        }

        /// <summary>
        /// Clear all override content types.
        /// </summary>
        public void ClearOverrideContentTypes()
        {
            if (this.overrideContentType != null)
                this.overrideContentType.Clear();
        }

        /// <summary>
        /// Parse the content types part.
        /// </summary>
        /// <exception cref="InvalidFormatException">
        /// Throws if the content type doesn't exist or the XML format is
        /// invalid.
        /// </exception>
        private void ParseContentTypesFile(Stream in1)
        {
            try
            {

                //in1.Position = 0;
                XPathDocument xpathdoc = DocumentHelper.ReadDocument(in1);
                XPathNavigator xpathnav = xpathdoc.CreateNavigator();
                XmlNamespaceManager nsMgr = new XmlNamespaceManager(xpathnav.NameTable);
                nsMgr.AddNamespace("x", TYPES_NAMESPACE_URI);

                XPathNodeIterator iterator = xpathnav.Select("//x:"+DEFAULT_TAG_NAME,nsMgr);
                while (iterator.MoveNext())
                {
                    // Default content types
                    //iterator.Current;
                    String extension = iterator.Current.GetAttribute(EXTENSION_ATTRIBUTE_NAME, xpathnav.NamespaceURI);
                    String contentType = iterator.Current.GetAttribute(CONTENT_TYPE_ATTRIBUTE_NAME, xpathnav.NamespaceURI);
                    AddDefaultContentType(extension, contentType);
                }
                iterator = xpathnav.Select("//x:" + OVERRIDE_TAG_NAME, nsMgr);

                while (iterator.MoveNext())
                {

                    // Overriden content types
                    //iterator.Current.MoveToNext();
                    Uri uri = PackagingUriHelper.ParseUri(iterator.Current.GetAttribute(PART_NAME_ATTRIBUTE_NAME, xpathnav.NamespaceURI), UriKind.RelativeOrAbsolute);
                    PackagePartName partName = PackagingUriHelper
                            .CreatePartName(uri);
                    String contentType = iterator.Current.GetAttribute(CONTENT_TYPE_ATTRIBUTE_NAME, xpathnav.NamespaceURI);
                    AddOverrideContentType(partName, contentType);
                }
            }
            catch (UriFormatException urie)
            {
                throw new InvalidFormatException(urie.Message);
            }
        }

        /// <summary>
        /// Save the contents type part.
        /// </summary>
        /// <param name="outStream">
        /// The output stream use to save the XML content of the content
        /// types part.
        /// </param>
        /// <returns><b>true</b> if the operation success, else <b>false</b>.</returns>
        public bool Save(Stream outStream)
        {
            XmlDocument xmlOutDoc = new XmlDocument();
            XmlNamespaceManager xmlnm = new XmlNamespaceManager(xmlOutDoc.NameTable);
            xmlnm.AddNamespace("x", TYPES_NAMESPACE_URI);
            XmlElement typesElem = xmlOutDoc.CreateElement(TYPES_TAG_NAME, TYPES_NAMESPACE_URI);
            xmlOutDoc.AppendChild(typesElem);

            // Adding default types
            IEnumerator<KeyValuePair<string, string>> contentTypes = defaultContentType.GetEnumerator();
            while (contentTypes.MoveNext())
            {
                AppendDefaultType(xmlOutDoc, typesElem, contentTypes.Current);
            }

            // Adding specific types if any exist
            if (overrideContentType != null)
            {

                IEnumerator<KeyValuePair<PackagePartName, string>> overrideContentTypes = overrideContentType.GetEnumerator();
                while (overrideContentTypes.MoveNext())
                {
                    AppendSpecificTypes(xmlOutDoc, typesElem, overrideContentTypes.Current);
                }
            }

            xmlOutDoc.Normalize();

            // Save content in the specified output stream
            return this.SaveImpl(xmlOutDoc, outStream);

        }

        /// <summary>
        /// Use to Append specific type XML elements, use by the save() method.
        /// </summary>
        /// <param name="root">
        /// XML parent element use to Append this override type element.
        /// </param>
        /// <param name="entry">
        /// The values to Append.
        /// </param>
        /// @see #save(java.io.OutputStream)
        private void AppendSpecificTypes(XmlDocument xmldoc, XmlElement root,
                KeyValuePair<PackagePartName, String> entry)
        {
            XmlElement elem = xmldoc.CreateElement(OVERRIDE_TAG_NAME, PackageNamespaces.CONTENT_TYPES);
            root.AppendChild(elem);
            elem.SetAttribute(
                    PART_NAME_ATTRIBUTE_NAME,
                    ((PackagePartName)entry.Key).Name);
            elem.SetAttribute(
                    CONTENT_TYPE_ATTRIBUTE_NAME, entry.Value);
        }

        /// <summary>
        /// Use to Append default types XML elements, use by the save() metid.
        /// </summary>
        /// <param name="root">
        /// XML parent element use to Append this default type element.
        /// </param>
        /// <param name="entry">
        /// The values to Append.
        /// </param>
        /// @see #save(java.io.OutputStream)
        private void AppendDefaultType(XmlDocument xmldoc, XmlElement root, KeyValuePair<String, String> entry)
        {
            XmlElement elem = xmldoc.CreateElement(DEFAULT_TAG_NAME,PackageNamespaces.CONTENT_TYPES);
            root.AppendChild(elem);
            elem.SetAttribute(EXTENSION_ATTRIBUTE_NAME, entry.Key);
            elem.SetAttribute(CONTENT_TYPE_ATTRIBUTE_NAME, entry.Value);
        }

        /// <summary>
        /// Specific implementation of the save method. Call by the save() method,
        /// call before exiting.
        /// </summary>
        /// <param name="out">
        /// The output stream use to write the content type XML.
        /// </param>
        public abstract bool SaveImpl(XmlDocument content, Stream out1);
    }

}