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
        public CoreProperties GetCoreProperties()
        {
            return core;
        }

        /**
         * Returns the extended document properties
         */
        public ExtendedProperties GetExtendedProperties()
        {
            return ext;
        }

        /**
         * Returns the custom document properties
         */
        public CustomProperties GetCustomProperties()
        {
            return cust;
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

            public String GetCategory()
            {
                return part.GetCategoryProperty();
            }
            public void SetCategory(String category)
            {
                part.SetCategoryProperty(category);
            }
            public String GetContentStatus()
            {
                return part.GetContentStatusProperty();
            }
            public void SetContentStatus(String contentStatus)
            {
                part.SetContentStatusProperty(contentStatus);
            }
            public String GetContentType()
            {
                return part.GetContentTypeProperty();
            }
            public void SetContentType(String contentType)
            {
                part.SetContentTypeProperty(contentType);
            }
            public DateTime? GetCreated()
            {
                return part.GetCreatedProperty();
            }
            public void SetCreated(Nullable<DateTime> date)
            {
                part.SetCreatedProperty(date);
            }
            public void SetCreated(String date)
            {
                part.SetCreatedProperty(date);
            }
            public String GetCreator()
            {
                return part.GetCreatorProperty();
            }
            public void SetCreator(String creator)
            {
                part.SetCreatorProperty(creator);
            }
            public String GetDescription()
            {
                return part.GetDescriptionProperty();
            }
            public void SetDescription(String description)
            {
                part.SetDescriptionProperty(description);
            }
            public String GetIdentifier()
            {
                return part.GetIdentifierProperty();
            }
            public void SetIdentifier(String identifier)
            {
                part.SetIdentifierProperty(identifier);
            }
            public String GetKeywords()
            {
                return part.GetKeywordsProperty();
            }
            public void SetKeywords(String keywords)
            {
                part.SetKeywordsProperty(keywords);
            }
            public DateTime? GetLastPrinted()
            {
                return part.GetLastPrintedProperty();
            }
            public void SetLastPrinted(Nullable<DateTime> date)
            {
                part.SetLastPrintedProperty(date);
            }
            public void SetLastPrinted(String date)
            {
                part.SetLastPrintedProperty(date);
            }
            public DateTime? GetModified()
            {
                return part.GetModifiedProperty();
            }
            public void SetModified(Nullable<DateTime> date)
            {
                part.SetModifiedProperty(date);
            }
            public void SetModified(String date)
            {
                part.SetModifiedProperty(date);
            }
            public String GetSubject()
            {
                return part.GetSubjectProperty();
            }
            public void SetSubjectProperty(String subject)
            {
                part.SetSubjectProperty(subject);
            }
            public void SetTitle(String title)
            {
                part.SetTitleProperty(title);
            }
            public String GetTitle()
            {
                return part.GetTitleProperty();
            }
            public String GetRevision()
            {
                return part.GetRevisionProperty();
            }
            public void SetRevision(String revision)
            {
                try
                {
                    long.Parse(revision);
                    part.SetRevisionProperty(revision);
                }
                catch (FormatException) { }
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
        }

    }

}
