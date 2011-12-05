/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

namespace NPOI
{

    using System.Text;
    using NPOI.OpenXml4Net.OPC.Internal;
    using System;
    using System.Collections.Generic;
    using OpenXmlFormats;

    /**
     * A {@link POITextExtractor} for returning the textual
     *  content of the OOXML file properties, eg author
     *  and title.
     */
    public class POIXMLPropertiesTextExtractor : POIXMLTextExtractor
    {
        /**
         * Creates a new POIXMLPropertiesTextExtractor for the
         *  given open document.
         */
        public POIXMLPropertiesTextExtractor(POIXMLDocument doc)
            : base(doc)
        {

        }
        /**
         * Creates a new POIXMLPropertiesTextExtractor, for the
         *  same file that another TextExtractor is already
         *  working on.
         */
        public POIXMLPropertiesTextExtractor(POIXMLTextExtractor otherExtractor)
            : base(otherExtractor.Document)
        {

        }

        private void AppendIfPresent(StringBuilder text, String thing, bool value)
        {
            AppendIfPresent(text, thing, value.ToString());
        }
        private void AppendIfPresent(StringBuilder text, String thing, int value)
        {
            AppendIfPresent(text, thing, value.ToString());
        }
        private void AppendIfPresent(StringBuilder text, String thing, DateTime? value)
        {
            if (value == null) { return; }
            AppendIfPresent(text, thing, value.ToString());
        }
        private void AppendIfPresent(StringBuilder text, String thing, String value)
        {
            if (value == null) { return; }
            text.Append(thing);
            text.Append(" = ");
            text.Append(value);
            text.Append("\n");
        }

        /**
         * Returns the core document properties, eg author
         */
        public String GetCorePropertiesText()
        {
            StringBuilder text = new StringBuilder();
            PackagePropertiesPart props =
                Document.GetProperties().GetCoreProperties().GetUnderlyingProperties();

            AppendIfPresent(text, "Category", props.GetCategoryProperty());
            AppendIfPresent(text, "Category", props.GetCategoryProperty());
            AppendIfPresent(text, "ContentStatus", props.GetContentStatusProperty());
            AppendIfPresent(text, "ContentType", props.GetContentTypeProperty());
            AppendIfPresent(text, "Created", props.GetCreatedProperty().Value);
            AppendIfPresent(text, "CreatedString", props.GetCreatedPropertyString());
            AppendIfPresent(text, "Creator", props.GetCreatorProperty());
            AppendIfPresent(text, "Description", props.GetDescriptionProperty());
            AppendIfPresent(text, "Identifier", props.GetIdentifierProperty());
            AppendIfPresent(text, "Keywords", props.GetKeywordsProperty());
            AppendIfPresent(text, "Language", props.GetLanguageProperty());
            AppendIfPresent(text, "LastModifiedBy", props.GetLastModifiedByProperty());
            AppendIfPresent(text, "LastPrinted", props.GetLastPrintedProperty());
            AppendIfPresent(text, "LastPrintedString", props.GetLastPrintedPropertyString());
            AppendIfPresent(text, "Modified", props.GetModifiedProperty());
            AppendIfPresent(text, "ModifiedString", props.GetModifiedPropertyString());
            AppendIfPresent(text, "Revision", props.GetRevisionProperty());
            AppendIfPresent(text, "Subject", props.GetSubjectProperty());
            AppendIfPresent(text, "Title", props.GetTitleProperty());
            AppendIfPresent(text, "Version", props.GetVersionProperty());

            return text.ToString();
        }
        /**
         * Returns the extended document properties, eg
         *  application
         */
        public String GetExtendedPropertiesText()
        {
            StringBuilder text = new StringBuilder();
            CT_Properties
                props = Document.GetProperties().GetExtendedProperties().GetUnderlyingProperties();

            AppendIfPresent(text, "Application", props.Application);
            AppendIfPresent(text, "AppVersion", props.AppVersion);
            AppendIfPresent(text, "Characters", props.Characters);
            AppendIfPresent(text, "CharactersWithSpaces", props.CharactersWithSpaces);
            AppendIfPresent(text, "Company", props.Company);
            AppendIfPresent(text, "HyperlinkBase", props.HyperlinkBase);
            AppendIfPresent(text, "HyperlinksChanged", props.HyperlinksChanged);
            AppendIfPresent(text, "Lines", props.Lines);
            AppendIfPresent(text, "LinksUpToDate", props.LinksUpToDate);
            AppendIfPresent(text, "Manager", props.Manager);
            AppendIfPresent(text, "Pages", props.Pages);
            AppendIfPresent(text, "Paragraphs", props.Paragraphs);
            AppendIfPresent(text, "PresentationFormat", props.PresentationFormat);
            AppendIfPresent(text, "Template", props.Template);
            AppendIfPresent(text, "TotalTime", props.TotalTime);

            return text.ToString();
        }
        /**
         * Returns the custom document properties, if
         *  there are any
         */
        public String GetCustomPropertiesText()
        {
            StringBuilder text = new StringBuilder();
            CT_Properties
                props = Document.GetProperties().GetCustomProperties().GetUnderlyingProperties();

            List<CT_Property> properties = props.GetPropertyList();
            for (int i = 0; i < properties.Count; i++)
            {
                // TODO - finish off
                String val = "(not implemented!)";

                text.Append(
                        properties[i].name +
                        " = " + val + "\n"
                );
            }

            return text.ToString();
        }

        public override String Text
        {
            get
            {
                return
                    GetCorePropertiesText() +
                    GetExtendedPropertiesText() +
                    GetCustomPropertiesText();
            }
        }

        public override POITextExtractor MetadataTextExtractor
        {
            get
            {
                throw new InvalidOperationException("You already have the Metadata Text Extractor, not recursing!");
            }
        }
    }





}

