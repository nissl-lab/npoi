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
    using NPOI.OpenXmlFormats;

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
            if (Document == null)
            {  // event based extractor does not have a document
                return "";
            }
            StringBuilder text = new StringBuilder();
            PackagePropertiesPart props =
                Document.GetProperties().CoreProperties.GetUnderlyingProperties();

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
            if (Document == null)
            {  // event based extractor does not have a document
                return "";
            }
            StringBuilder text = new StringBuilder();
            CT_ExtendedProperties
                props = Document.GetProperties().ExtendedProperties.GetUnderlyingProperties();

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
            if (Document == null)
            {  // event based extractor does not have a document
                return "";
            }
            StringBuilder text = new StringBuilder();
            CT_CustomProperties props = Document.GetProperties().CustomProperties.GetUnderlyingProperties();

            List<CT_Property> properties = props.GetPropertyList();
            foreach (CT_Property property in properties)
            {
                String val = "(not implemented!)";
                //val = property.Item.ToString();
                //if (property.IsSetLpwstr())
                //{
                //    val = property.GetLpwstr();
                //}
                //else if (property.IsSetLpstr())
                //{
                //    val = property.GetLpstr();
                //}
                //else if (property.IsSetDate())
                //{
                //    val = property.GetDate().toString();
                //}
                //else if (property.IsSetFiletime())
                //{
                //    val = property.GetFiletime().toString();
                //}
                //else if (property.IsSetBool())
                //{
                //    val = Boolean.toString(property.GetBool());
                //}

                //// Integers
                //else if (property.IsSetI1())
                //{
                //    val = Integer.toString(property.GetI1());
                //}
                //else if (property.IsSetI2())
                //{
                //    val = Integer.toString(property.GetI2());
                //}
                //else if (property.IsSetI4())
                //{
                //    val = Integer.toString(property.GetI4());
                //}
                //else if (property.IsSetI8())
                //{
                //    val = Long.toString(property.GetI8());
                //}
                //else if (property.IsSetInt())
                //{
                //    val = Integer.toString(property.GetInt());
                //}

                //// Unsigned Integers
                //else if (property.IsSetUi1())
                //{
                //    val = Integer.toString(property.GetUi1());
                //}
                //else if (property.IsSetUi2())
                //{
                //    val = Integer.toString(property.GetUi2());
                //}
                //else if (property.IsSetUi4())
                //{
                //    val = Long.toString(property.GetUi4());
                //}
                //else if (property.IsSetUi8())
                //{
                //    val = property.GetUi8().toString();
                //}
                //else if (property.IsSetUint())
                //{
                //    val = Long.toString(property.GetUint());
                //}

                //// Reals
                //else if (property.IsSetR4())
                //{
                //    val = Float.toString(property.GetR4());
                //}
                //else if (property.IsSetR8())
                //{
                //    val = Double.toString(property.GetR8());
                //}
                //else if (property.IsSetDecimal())
                //{
                //    BigDecimal d = property.GetDecimal();
                //    if (d == null)
                //    {
                //        val = null;
                //    }
                //    else
                //    {
                //        val = d.toPlainString();
                //    }
                //}

                //else if (property.IsSetArray())
                //{
                //    // TODO Fetch the array values and output 
                //}
                //else if (property.IsSetVector())
                //{
                //    // TODO Fetch the vector values and output
                //}

                //else if (property.IsSetBlob() || property.IsSetOblob())
                //{
                //    // TODO Decode, if possible
                //}
                //else if (property.IsSetStream() || property.IsSetOstream() ||
                //         property.IsSetVstream())
                //{
                //    // TODO Decode, if possible
                //}
                //else if (property.IsSetStorage() || property.IsSetOstorage())
                //{
                //    // TODO Decode, if possible
                //}

                text.Append(
                      property.name +
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

