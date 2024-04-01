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
using System.IO;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.Util;
using System.Globalization;
using System.Text.RegularExpressions;
using NPOI.Util;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /// <summary>
    /// Represents the core properties part of a package.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public class PackagePropertiesPart : PackagePart, PackageProperties
    {
        public static String NAMESPACE_DC = "http://purl.org/dc/elements/1.1/";

        public static String NAMESPACE_DC_URI = "http://purl.org/dc/elements/1.1/";

        public static String NAMESPACE_CP_URI = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

        public static String NAMESPACE_DCTERMS_URI = "http://purl.org/dc/terms/";

        public static String NAMESPACE_XSI_URI = "http://www.w3.org/2001/XMLSchema-instance";
        private static String DEFAULT_DATEFORMAT = "yyyy-MM-dd'T'HH:mm:ss'Z'";
        //private static String ALTERNATIVE_DATEFORMAT = "yyyy-MM-dd'T'HH:mm:ss.SS'Z'";
        private static String[] DATE_FORMATS = new String[]{
            DEFAULT_DATEFORMAT,
            "yyyy-MM-dd'T'HH:mm:ss.ff'Z'",
            "yyyy-MM-dd",
        };

        //Had to add this and TIME_ZONE_PAT to handle tz with colons.
        //When we move to Java 7, we should be able to add another
        //date format to DATE_FORMATS that uses XXX and get rid of this
        //and TIME_ZONE_PAT
        private String[] TZ_DATE_FORMATS = new String[]{
            "yyyy-MM-dd'T'HH:mm:ssz",
            "yyyy-MM-dd'T'HH:mm:ss.fz",
            "yyyy-MM-dd'T'HH:mm:ss.ffz",
            "yyyy-MM-dd'T'HH:mm:ss.fffz",
        };

        private Regex TIME_ZONE_PAT = new Regex("([-+]\\d\\d):?(\\d\\d)");

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="pack">
        /// Container package.
        /// </param>
        /// <param name="partName">
        /// Name of this part.
        /// </param>
        /// <exception cref="InvalidFormatException">InvalidFormatException
        /// Throws if the content is invalid.
        /// </exception>
        public PackagePropertiesPart(OPCPackage pack, PackagePartName partName)
            : base(pack, partName, ContentTypes.CORE_PROPERTIES_PART)
        {

        }

        /// <summary>
        /// <para>
        /// A categorization of the content of this package.
        /// </para>
        /// <para>
        /// [Example: Example values for this property might include: Resume, Letter,
        /// Financial Forecast, Proposal, Technical Presentation, and so on. This
        /// value might be used by an application's user interface to facilitate
        /// navigation of a large Set of documents. end example]
        /// </para>
        /// </summary>
        protected String category = null;

        /// <summary>
        /// <para>
        /// The status of the content.
        /// </para>
        /// <para>
        /// [Example: Values might include "Draft", "Reviewed", and "Final". end
        /// example]
        /// </para>
        /// </summary>
        protected String contentStatus = null;

        /// <summary>
        /// <para>
        /// The type of content represented, generally defined by a specific use and
        /// intended audience.
        /// </para>
        /// <para>
        /// [Example: Values might include "Whitepaper", "Security Bulletin", and
        /// "Exam". end example] [Note: This property is distinct from MIME content
        /// types as defined in RFC 2616. end note]
        /// </para>
        /// </summary>
        protected String contentType = null;

        /// <summary>
        /// Date of creation of the resource.
        /// </summary>
        protected Nullable<DateTime> created = new Nullable<DateTime>();

        /// <summary>
        /// An entity primarily responsible for making the content of the resource.
        /// </summary>
        protected String creator = null;

        /// <summary>
        /// <para>
        /// An explanation of the content of the resource.
        /// </para>
        /// <para>
        /// [Example: Values might include an abstract, table of contents, reference
        /// to a graphical representation of content, and a free-text account of the
        /// content. end example]
        /// </para>
        /// </summary>
        protected String description = null;

        /// <summary>
        /// An unambiguous reference to the resource within a given context.
        /// </summary>
        protected String identifier = null;

        /// <summary>
        /// A delimited Set of keywords to support searching and indexing. This is
        /// typically a list of terms that are not available elsewhere in the
        /// properties.
        /// </summary>
        protected String keywords = null;

        /// <summary>
        /// <para>
        /// The language of the intellectual content of the resource.
        /// </para>
        /// <para>
        /// [Note: IETF RFC 3066 provides guidance on encoding to represent
        /// languages. end note]
        /// </para>
        /// </summary>
        protected String language = null;

        /// <summary>
        /// <para>
        /// The user who performed the last modification. The identification is
        /// environment-specific.
        /// </para>
        /// <para>
        /// [Example: A name, email address, or employee ID. end example] It is
        /// recommended that this value be as concise as possible.
        /// </para>
        /// </summary>
        protected String lastModifiedBy = null;

        /// <summary>
        /// The date and time of the last printing.
        /// </summary>
        protected Nullable<DateTime> lastPrinted = new Nullable<DateTime>();

        /// <summary>
        /// Date on which the resource was changed.
        /// </summary>
        protected Nullable<DateTime> modified = new Nullable<DateTime>();

        /// <summary>
        /// <para>
        /// The revision number.
        /// </para>
        /// <para>
        /// [Example: This value might indicate the number of saves or revisions,
        /// provided the application updates it after each revision. end example]
        /// </para>
        /// </summary>
        protected String revision = null;

        /// <summary>
        /// The topic of the content of the resource.
        /// </summary>
        protected String subject = null;

        /// <summary>
        /// The name given to the resource.
        /// </summary>
        protected String title = null;

        /// <summary>
        /// The version number. This value is Set by the user or by the application.
        /// </summary>
        protected String version = null;

        /*
         * Getters and Setters
         */

        /// <summary>
        /// <para>
        /// Get the category property.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getCategoryProperty()" />
        /// </para>
        /// </summary>
        public String GetCategoryProperty() {
            return category;
        }

        /// <summary>
        /// <para>
        /// Get content status.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getContentStatusProperty()" />
        /// </para>
        /// </summary>
        public String GetContentStatusProperty() {
            return contentStatus;
        }

        /// <summary>
        /// <para>
        /// Get content type.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getContentTypeProperty()" />
        /// </para>
        /// </summary>
        public String GetContentTypeProperty() {
            return contentType;
        }

        /// <summary>
        /// <para>
        /// Get created date.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getCreatedProperty()" />
        /// </para>
        /// </summary>
        public Nullable<DateTime> GetCreatedProperty() {
            return created;
        }

        /// <summary>
        /// Get created date formated into a String.
        /// </summary>
        /// <returns>A string representation of the created date.</returns>
        public String GetCreatedPropertyString() {
            return GetDateValue(created);
        }

        /// <summary>
        /// <para>
        /// Get creator.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getCreatorProperty()" />
        /// </para>
        /// </summary>
        public String GetCreatorProperty() {
            return creator;
        }

        /// <summary>
        /// <para>
        /// Get description.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getDescriptionProperty()" />
        /// </para>
        /// </summary>
        public String GetDescriptionProperty() {
            return description;
        }

        /// <summary>
        /// <para>
        /// Get identifier.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getIdentifierProperty()" />
        /// </para>
        /// </summary>
        public String GetIdentifierProperty() {
            return identifier;
        }

        /// <summary>
        /// <para>
        /// Get keywords.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getKeywordsProperty()" />
        /// </para>
        /// </summary>
        public String GetKeywordsProperty() {
            return keywords;
        }

        /// <summary>
        /// <para>
        /// Get the language.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getLanguageProperty()" />
        /// </para>
        /// </summary>
        public String GetLanguageProperty() {
            return language;
        }

        /// <summary>
        /// <para>
        /// Get the author of last modifications.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getLastModifiedByProperty()" />
        /// </para>
        /// </summary>
        public String GetLastModifiedByProperty() {
            return lastModifiedBy;
        }

        /// <summary>
        /// <para>
        /// Get last printed date.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getLastPrintedProperty()" />
        /// </para>
        /// </summary>
        public Nullable<DateTime> GetLastPrintedProperty() {
            return lastPrinted;
        }

        /// <summary>
        /// Get last printed date formated into a String.
        /// </summary>
        /// <returns>A string representation of the last printed date.</returns>
        public String GetLastPrintedPropertyString() {
            return GetDateValue(lastPrinted);
        }

        /// <summary>
        /// <para>
        /// Get modified date.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getModifiedProperty()" />
        /// </para>
        /// </summary>
        public Nullable<DateTime> GetModifiedProperty() {
            return modified;
        }

        /// <summary>
        /// Get modified date formated into a String.
        /// </summary>
        /// <returns>A string representation of the modified date.</returns>
        public String GetModifiedPropertyString() {
            if (modified == null)
                return GetDateValue(new Nullable<DateTime>(new DateTime()));
            else
                return GetDateValue(modified);
        }

        /// <summary>
        /// <para>
        /// Get revision.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getRevisionProperty()" />
        /// </para>
        /// </summary>
        public String GetRevisionProperty() {
            return revision;
        }

        /// <summary>
        /// <para>
        /// Get subject.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getSubjectProperty()" />
        /// </para>
        /// </summary>
        public String GetSubjectProperty() {
            return subject;
        }

        /// <summary>
        /// <para>
        /// Get title.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getTitleProperty()" />
        /// </para>
        /// </summary>
        public String GetTitleProperty() {
            return title;
        }

        /// <summary>
        /// <para>
        /// Get version.
        /// </para>
        /// <para>
        /// <see cref="org.apache.poi.OpenXml4Net.opc.PackageProperties.getVersionProperty()" />
        /// </para>
        /// </summary>
        public String GetVersionProperty() {
            return version;
        }

        /// <summary>
        /// <para>
        /// Set the category.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetCategoryProperty(String)" />
        /// </para>
        /// </summary>
        public void SetCategoryProperty(String category) {
            this.category = SetStringValue(category);
        }

        /// <summary>
        /// <para>
        /// Set the content status.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetContentStatusProperty(String)" />
        /// </para>
        /// </summary>
        public void SetContentStatusProperty(String contentStatus) {
            this.contentStatus = SetStringValue(contentStatus);
        }

        /// <summary>
        /// <para>
        /// Set the content type.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetContentTypeProperty(String)" />
        /// </para>
        /// </summary>
        public void SetContentTypeProperty(String contentType) {
            this.contentType = SetStringValue(contentType);
        }

        /// <summary>
        /// <para>
        /// Set the created date.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetCreatedProperty(org.apache.poi.OpenXml4Net.util.Nullable)" />
        /// </para>
        /// </summary>
        public void SetCreatedProperty(String created) {
            try {
                this.created = SetDateValue(created);
            } catch (InvalidFormatException e) {
                throw new ArgumentException("Date for created could not be parsed: " + created, e);
            }
        }

        /// <summary>
        /// <para>
        /// Set the created date.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetCreatedProperty(org.apache.poi.OpenXml4Net.util.Nullable)" />
        /// </para>
        /// </summary>
        public void SetCreatedProperty(Nullable<DateTime> created) {
            if (created != null)
                this.created = created;
        }

        /// <summary>
        /// <para>
        /// Set the creator.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetCreatorProperty(String)" />
        /// </para>
        /// </summary>
        public void SetCreatorProperty(String creator) {
            this.creator = SetStringValue(creator);
        }

        /// <summary>
        /// <para>
        /// Set the description.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetDescriptionProperty(String)" />
        /// </para>
        /// </summary>
        public void SetDescriptionProperty(String description) {
            this.description = SetStringValue(description);
        }

        /// <summary>
        /// <para>
        /// Set identifier.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetIdentifierProperty(String)" />
        /// </para>
        /// </summary>
        public void SetIdentifierProperty(String identifier) {
            this.identifier = SetStringValue(identifier);
        }

        /// <summary>
        /// <para>
        /// Set keywords.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetKeywordsProperty(String)" />
        /// </para>
        /// </summary>
        public void SetKeywordsProperty(String keywords) {
            this.keywords = SetStringValue(keywords);
        }

        /// <summary>
        /// <para>
        /// Set language.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetLanguageProperty(String)" />
        /// </para>
        /// </summary>
        public void SetLanguageProperty(String language) {
            this.language = SetStringValue(language);
        }

        /// <summary>
        /// <para>
        /// Set last modifications author.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetLastModifiedByProperty(String)" />
        /// </para>
        /// </summary>
        public void SetLastModifiedByProperty(String lastModifiedBy) {
            this.lastModifiedBy = SetStringValue(lastModifiedBy);
        }

        /// <summary>
        /// <para>
        /// Set last printed date.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetLastPrintedProperty(org.apache.poi.OpenXml4Net.util.Nullable)" />
        /// </para>
        /// </summary>
        public void SetLastPrintedProperty(String lastPrinted) {
            try {
                this.lastPrinted = SetDateValue(lastPrinted);
            } catch (InvalidFormatException e) {
                new ArgumentException("lastPrinted  : "
                        + e.Message, e);
            }
        }

        /// <summary>
        /// <para>
        /// Set last printed date.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetLastPrintedProperty(org.apache.poi.OpenXml4Net.util.Nullable)" />
        /// </para>
        /// </summary>
        public void SetLastPrintedProperty(Nullable<DateTime> lastPrinted) {
            if (lastPrinted != null)
                this.lastPrinted = lastPrinted;
        }

        /// <summary>
        /// <para>
        /// Set last modification date.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetModifiedProperty(org.apache.poi.OpenXml4Net.util.Nullable)" />
        /// </para>
        /// </summary>
        public void SetModifiedProperty(String modified) {
            try {
                this.modified = SetDateValue(modified);
            } catch (InvalidFormatException e) {
                new ArgumentException("modified  : "
                        + e.Message, e);
            }
        }

        /// <summary>
        /// <para>
        /// Set last modification date.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetModifiedProperty(org.apache.poi.OpenXml4Net.util.Nullable)" />
        /// </para>
        /// </summary>
        public void SetModifiedProperty(Nullable<DateTime> modified) {
            if (modified.HasValue)
                this.modified = modified;
        }

        /// <summary>
        /// <para>
        /// Set revision.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetRevisionProperty(String)" />
        /// </para>
        /// </summary>
        public void SetRevisionProperty(String revision) {
            this.revision = SetStringValue(revision);
        }

        /// <summary>
        /// <para>
        /// Set subject.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetSubjectProperty(String)" />
        /// </para>
        /// </summary>
        public void SetSubjectProperty(String subject) {
            this.subject = SetStringValue(subject);
        }

        /// <summary>
        /// <para>
        /// Set title.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetTitleProperty(String)" />
        /// </para>
        /// </summary>
        public void SetTitleProperty(String title) {
            this.title = SetStringValue(title);
        }

        /// <summary>
        /// <para>
        /// Set version.
        /// </para>
        /// <para>
        /// <see cref="PackageProperties.SetVersionProperty(String)" />
        /// </para>
        /// </summary>
        public void SetVersionProperty(String version) {
            this.version = SetStringValue(version);
        }

        /// <summary>
        /// Convert a strig value into a String
        /// </summary>
        private String SetStringValue(String s) {
            if (s == null || s.Equals(""))
                return null;
            else
                return s;
        }

        /// <summary>
        /// Convert a string value represented a date into a DateTime?.
        /// </summary>
        /// <exception cref="InvalidFormatException">InvalidFormatException
        /// Throws if the date format isnot valid.
        /// </exception>
        private DateTime? SetDateValue(String dateStr) {
            if (dateStr == null || dateStr.Equals(""))
            {
                return new Nullable<DateTime>();
            }
            Match m = TIME_ZONE_PAT.Match(dateStr);
            String dateTzStr;
            if (m.Success)
            {
                dateTzStr = dateStr.Substring(0, m.Index) +
                        m.Groups[1].Value + m.Groups[2].Value;
                foreach (String fStr in TZ_DATE_FORMATS)
                {
                    SimpleDateFormat df = new SimpleDateFormat(fStr);
                    df.TimeZone = TimeZoneInfo.Utc;
                    try
                    {
                        return df.Parse(dateTzStr);
                    }
                    catch (FormatException)
                    {
                    }
                }
            }
            dateTzStr = dateStr.EndsWith("Z") ? dateStr : (dateStr + "Z");
            foreach (String fStr in DATE_FORMATS)
            {
                SimpleDateFormat df = new SimpleDateFormat(fStr);
                df.TimeZone = TimeZoneInfo.Utc;
                try
                {
                    return df.Parse(dateTzStr).ToUniversalTime();
                }
                catch (FormatException)
                {
                }
            }
            //if you're here, no pattern matched, throw exception
            StringBuilder sb = new StringBuilder();
            int i = 0;
            foreach (String fStr in TZ_DATE_FORMATS)
            {
                if (i++ > 0)
                {
                    sb.Append(", ");
                }
                sb.Append(fStr);
            }
            foreach (String fStr in DATE_FORMATS)
            {
                sb.Append(", ").Append(fStr);
            }
            throw new InvalidFormatException("Date " + dateStr + " not well formatted, "
                    + "expected format in: " + sb.ToString());

        }

        /// <summary>
        /// Convert a DateTime? into a String.
        /// </summary>
        /// <param name="d">
        /// The Date to convert.
        /// </param>
        /// <returns>The formated date or null.</returns>
        /// @see java.util.SimpleDateFormat
        private String GetDateValue(DateTime? d) {
            if (!d.HasValue || d == null || d.Equals(""))
            {
                return "";
            }

            SimpleDateFormat df = new SimpleDateFormat(DEFAULT_DATEFORMAT);
            df.TimeZone = TimeZoneInfo.Utc;
            return df.Format(d.Value, CultureInfo.InvariantCulture);
        }


        protected override Stream GetInputStreamImpl() {
            throw new InvalidOperationException("Operation not authorized");
        }

        protected override Stream GetOutputStreamImpl()
        {
            throw new InvalidOperationException("Operation not authorized");
        }


        public override bool Save(Stream zos) {
            throw new InvalidOperationException("Operation not authorized");
        }


        public override bool Load(Stream ios) {
            throw new InvalidOperationException("Operation not authorized");
        }


        public override void Close() {
            // Do nothing
        }


        public override void Flush() {
            // Do nothing
        }
    }

}
