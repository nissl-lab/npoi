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

using System;
using System.IO;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.OpenXmlFormats;

namespace NPOI
{
    /**
 * The core document properties
 */
    public class CoreProperties
    {
        private PackagePropertiesPart part;
        internal CoreProperties(PackagePropertiesPart part)
        {
            this.part = part;
        }

        public String Category
        {
            get
            {
                return part.GetCategoryProperty();
            }
            set 
            {
                part.SetCategoryProperty(value);
            }
        }
        public String ContentStatus
        {
            get
            {
                return part.GetContentStatusProperty();
            }
            set 
            {
                part.SetContentStatusProperty(value);
            }
        }
        public String ContentType
        {
            get
            {
                return part.GetContentTypeProperty();
            }
            set 
            {
                part.SetContentTypeProperty(value);
            }
        }
        public DateTime? Created
        {
            get
            {
                return part.GetCreatedProperty();
            }
            set 
            {
                part.SetCreatedProperty(value);    
            }
        }
        public void SetCreated(String date)
        {
            part.SetCreatedProperty(date);
        }
        public String Creator
        {
            get
            {
                return part.GetCreatorProperty();
            }
            set 
            {
                part.SetCreatorProperty(value);
            }
        }
        public String Description
        {
            get
            {
                return part.GetDescriptionProperty();
            }
            set 
            {
                part.SetDescriptionProperty(value);
            }
        }
        public String Identifier
        {
            get
            {
                return part.GetIdentifierProperty();
            }
            set 
            {
                part.SetIdentifierProperty(value);
            }
        }
        public String Keywords
        {
            get
            {
                return part.GetKeywordsProperty();
            }
            set 
            {
                part.SetKeywordsProperty(value);
            }
        }
        public DateTime? LastPrinted
        {
            get
            {
                return part.GetLastPrintedProperty();
            }
            set 
            {
                part.SetLastPrintedProperty(value);
            }
        }
        public void SetLastPrinted(String date)
        {
            part.SetLastPrintedProperty(date);
        }
        public DateTime? Modified
        {
            get
            {
                return part.GetModifiedProperty();
            }
            set 
            {
                part.SetModifiedProperty(value);
            }
        }
        public void SetModified(String date)
        {
            part.SetModifiedProperty(date);
        }
        public String Subject
        {
            get
            {
                return part.GetSubjectProperty();
            }
            set 
            {
                part.SetSubjectProperty(value);
            }
        }
        public String Title
        {
            get
            {
                return part.GetTitleProperty();
            }
            set 
            {
                part.SetTitleProperty(value);
            }
        }
        public String Revision
        {
            get
            {
                return part.GetRevisionProperty();
            }
            set
            {
                try
                {
                    long.Parse(value);
                    part.SetRevisionProperty(value);
                }
                catch (FormatException) { }            
            }
        }

        public PackagePropertiesPart GetUnderlyingProperties()
        {
            return part;
        }
    }

    /**
     * Extended document properties
     */
    public class ExtendedProperties
    {
        public ExtendedPropertiesDocument props;
        internal ExtendedProperties(ExtendedPropertiesDocument props)
        {
            this.props = props;
        }

        public CT_ExtendedProperties GetUnderlyingProperties()
        {
            return props.GetProperties();
        }

        public String Template
        {
            get
            {
                return props.GetProperties().Template;
            }
        }
        public String Manager
        {
            get { return props.GetProperties().Manager; }
        }
        public String Company
        {
            get { return props.GetProperties().Company; }
        }
        public String PresentationFormat
        {
            get { return props.GetProperties().PresentationFormat; }
        }
        public String Application
        {
            get { return props.GetProperties().Application; }
        }
        public String AppVersion
        {
            get { return props.GetProperties().AppVersion; }
        }

        public int Pages
        {
            get
            {
                if (props.GetProperties().IsSetPages())
                {
                    return props.GetProperties().Pages;
                }
                return -1;
            }
        }
        public int Words
        {
            get
            {
                if (props.GetProperties().IsSetWords())
                {
                    return props.GetProperties().Words;
                }
                return -1;
            }
        }
        public int Characters
        {
            get
            {
                if (props.GetProperties().IsSetCharacters())
                {
                    return props.GetProperties().Characters;
                }
                return -1;
            }
        }
        public int CharactersWithSpaces
        {
            get
            {
                if (props.GetProperties().IsSetCharactersWithSpaces())
                {
                    return props.GetProperties().CharactersWithSpaces;
                }
                return -1;
            }
        }
        public int Lines
        {
            get
            {
                if (props.GetProperties().IsSetLines())
                {
                    return props.GetProperties().Lines;
                }
                return -1;
            }
        }
        public int Paragraphs
        {
            get
            {
                if (props.GetProperties().IsSetParagraphs())
                {
                    return props.GetProperties().Paragraphs;
                }
                return -1;
            }
        }
        public int Slides
        {
            get
            {
                if (props.GetProperties().IsSetSlides())
                {
                    return props.GetProperties().Slides;
                }
                return -1;
            }
        }
        public int Notes
        {
            get
            {
                if (props.GetProperties().IsSetNotes())
                {
                    return props.GetProperties().Notes;
                }
                return -1;
            }
        }
        public int TotalTime
        {
            get
            {
                if (props.GetProperties().IsSetTotalTime())
                {
                    return props.GetProperties().TotalTime;
                }
                return -1;
            }
        }
        public int HiddenSlides
        {
            get
            {
                if (props.GetProperties().IsSetHiddenSlides())
                {
                    return props.GetProperties().HiddenSlides;
                }
                return -1;
            }
        }
        public int MMClips
        {
            get
            {
                if (props.GetProperties().IsSetMMClips())
                {
                    return props.GetProperties().MMClips;
                }
                return -1;
            }
        }

        public String HyperlinkBase
        {
            get { return props.GetProperties().HyperlinkBase; }
        }
    }

    /**
     *  Custom document properties
     */
    public class CustomProperties
    {
        /**
         *  Each custom property element Contains an fmtid attribute
         *  with the same GUID value ({D5CDD505-2E9C-101B-9397-08002B2CF9AE}).
         */
        public static String FORMAT_ID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

        public CustomPropertiesDocument props;
        internal CustomProperties(CustomPropertiesDocument props)
        {
            this.props = props;
        }

        public CT_CustomProperties GetUnderlyingProperties()
        {
            return props.GetProperties();
        }

        /**
         * Add a new property
         *
         * @param name the property name
         * @throws IllegalArgumentException if a property with this name already exists
         */
        private CT_Property Add(String name)
        {
            if (Contains(name))
            {
                throw new ArgumentException("A property with this name " +
                        "already exists in the custom properties");
            }

            CT_Property p = props.GetProperties().AddNewProperty();
            int pid = NextPid();
            p.pid = pid;
            p.fmtid = FORMAT_ID;
            p.name = name;
            return p;
        }

        /**
         * Add a new string property
         *
         * @throws IllegalArgumentException if a property with this name already exists
         */
        public void AddProperty(String name, String value)
        {
            CT_Property p = Add(name);
            p.ItemElementName = ItemChoiceType.lpwstr;
            p.Item = value;
        }

        /**
         * Add a new double property
         *
         * @throws IllegalArgumentException if a property with this name already exists
         */
        public void AddProperty(String name, double value)
        {
            CT_Property p = Add(name);
            p.ItemElementName = ItemChoiceType.r8;
            p.Item = value;
        }

        /**
         * Add a new integer property
         *
         * @throws IllegalArgumentException if a property with this name already exists
         */
        public void AddProperty(String name, int value)
        {
            CT_Property p = Add(name);
            p.ItemElementName = ItemChoiceType.i4;
            p.Item = value;
        }

        /**
         * Add a new bool property
         *
         * @throws IllegalArgumentException if a property with this name already exists
         */
        public void AddProperty(String name, bool value)
        {
            CT_Property p = Add(name);
            p.ItemElementName = ItemChoiceType.@bool;
            p.Item = value;
        }

        /**
         * Generate next id that uniquely relates a custom property
         *
         * @return next property id starting with 2
         */
        protected int NextPid()
        {
            int propid = 1;
            foreach (CT_Property p in props.GetProperties().GetPropertyList())
            {
                if (p.pid > propid) propid = p.pid;
            }
            return propid + 1;
        }

        /**
         * Check if a property with this name already exists in the collection of custom properties
         *
         * @param name the name to check
         * @return whether a property with the given name exists in the custom properties
         */
        public bool Contains(String name)
        {
            foreach (CT_Property p in props.GetProperties().GetPropertyList())
            {
                if (p.name.Equals(name)) return true;
            }
            return false;
        }

        /**
         * Retrieve the custom property with this name, or null if none exists.
         *
         * You will need to test the various isSetX methods to work out
         *  what the type of the property is, before fetching the 
         *  appropriate value for it.
         *
         * @param name the name of the property to fetch
         */
        public CT_Property GetProperty(String name) {
            foreach(CT_Property p in props.GetProperties().GetPropertyList()){
                if(p.name.Equals(name)) {
                    return p;
                }
            }
            return null;
        }
    }
    /**
     * Wrapper around the two different kinds of OOXML properties
     *  a document can have
     */
    public class POIXMLProperties
    {
        private OPCPackage pkg;
        private CoreProperties core;
        private ExtendedProperties ext;
        private CustomProperties cust;

        private PackagePart extPart;
        private PackagePart custPart;


        private static ExtendedPropertiesDocument NEW_EXT_INSTANCE;
        private static CustomPropertiesDocument NEW_CUST_INSTANCE;
        static POIXMLProperties()
        {
            NEW_EXT_INSTANCE = new ExtendedPropertiesDocument();
            NEW_EXT_INSTANCE.AddNewProperties();

            NEW_CUST_INSTANCE = new CustomPropertiesDocument();
            NEW_CUST_INSTANCE.AddNewProperties();
        }

        public POIXMLProperties(OPCPackage docPackage)
        {
            this.pkg = docPackage;

            // Core properties
            core = new CoreProperties((PackagePropertiesPart)pkg.GetPackageProperties());

            // Extended properties
            PackageRelationshipCollection extRel =
                pkg.GetRelationshipsByType(PackageRelationshipTypes.EXTENDED_PROPERTIES);
            if (extRel.Size == 1)
            {
                extPart = pkg.GetPart(extRel.GetRelationship(0));
                ExtendedPropertiesDocument props = ExtendedPropertiesDocument.Parse(
                     extPart.GetInputStream()
                );
                ext = new ExtendedProperties(props);
            }
            else
            {
                extPart = null;
                ext = new ExtendedProperties((ExtendedPropertiesDocument)NEW_EXT_INSTANCE.Copy());
            }

            // Custom properties
            PackageRelationshipCollection custRel =
                pkg.GetRelationshipsByType(PackageRelationshipTypes.CUSTOM_PROPERTIES);
            if (custRel.Size == 1)
            {
                custPart = pkg.GetPart(custRel.GetRelationship(0));
                CustomPropertiesDocument props = CustomPropertiesDocument.Parse(
                        custPart.GetInputStream()
                );
                cust = new CustomProperties(props);
            }
            else
            {
                custPart = null;
                cust = new CustomProperties((CustomPropertiesDocument)NEW_CUST_INSTANCE.Copy());
            }
        }

        /**
         * Returns the core document properties
         */
        public CoreProperties CoreProperties
        {
            get
            {
                return core;
            }
        }

        /**
         * Returns the extended document properties
         */
        public ExtendedProperties ExtendedProperties
        {
            get
            {
                return ext;
            }
        }

        /**
         * Returns the custom document properties
         */
        public CustomProperties CustomProperties
        {
            get
            {
                return cust;
            }
        }

        /**
         * Commit Changes to the underlying OPC namespace
         */
        public virtual void Commit()
        {

            if (extPart == null && !NEW_EXT_INSTANCE.ToString().Equals(ext.props.ToString()))
            {
                try
                {
                    PackagePartName prtname = PackagingUriHelper.CreatePartName("/docProps/app.xml");
                    pkg.AddRelationship(prtname, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
                    extPart = pkg.CreatePart(prtname, "application/vnd.openxmlformats-officedocument.extended-properties+xml");
                }
                catch (InvalidFormatException e)
                {
                    throw new POIXMLException(e);
                }
            }
            if (custPart == null && !NEW_CUST_INSTANCE.ToString().Equals(cust.props.ToString()))
            {
                try
                {
                    PackagePartName prtname = PackagingUriHelper.CreatePartName("/docProps/custom.xml");
                    pkg.AddRelationship(prtname, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
                    custPart = pkg.CreatePart(prtname, "application/vnd.openxmlformats-officedocument.custom-properties+xml");
                }
                catch (InvalidFormatException e)
                {
                    throw new POIXMLException(e);
                }
            }
            if (extPart != null)
            {
                Stream out1 = extPart.GetOutputStream();

                if (extPart.Size > 0)
                    extPart.Clear();
                ext.props.Save(out1);
                out1.Close();
            }
            if (custPart != null)
            {
                Stream out1 = custPart.GetOutputStream();
                cust.props.Save(out1);
                out1.Close();
            }


        }


    }

}
