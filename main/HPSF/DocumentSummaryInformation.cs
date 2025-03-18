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
using System.Collections;
using System.Collections.Generic;

namespace NPOI.HPSF
{


    using NPOI.HPSF.Wellknown;
    using NPOI.Util;
    using Org.BouncyCastle.Asn1.Ocsp;

    /// <summary>
    /// Convenience class representing a DocumentSummary Information stream in a
    /// Microsoft Office document.
    /// </summary>
    /// @see SummaryInformation
    public class DocumentSummaryInformation : SpecialPropertySet
    {
        /// <summary>
        /// The document name a document summary information stream
        /// usually has in a POIFS filesystem.
        /// </summary>
        public static String DEFAULT_STREAM_NAME =
            "\x0005DocumentSummaryInformation";
        public override PropertyIDMap PropertySetIDMap => PropertyIDMap.DocumentSummaryInformationProperties;


        /// <summary>
        /// Creates an empty {@link DocumentSummaryInformation}.
        /// </summary>
        public DocumentSummaryInformation()
        {
            FirstSection.SetFormatID(SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID[0]);
        }


        /// <summary>
        /// Creates a {@link DocumentSummaryInformation} from a given
        /// {@link PropertySet}.
        /// </summary>
        /// <param name="ps">A property Set which should be created from a
        /// document summary information stream.
        /// </param>
        /// <throws name="UnexpectedPropertySetTypeException">if <c>ps</c>
        /// does not contain a document summary information stream.
        /// </throws>
        public DocumentSummaryInformation(PropertySet ps)
        : base(ps)
        {

            ;
            if(!IsDocumentSummaryInformation)
            {
                throw new UnexpectedPropertySetTypeException("Not a " + GetType().Name);
            }
        }


        /// <summary>
        /// get or set the category (or <c>null</c>).
        /// </summary>
        /// <return>category value</return>
        public String Category
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_CATEGORY);
            set => FirstSection.SetProperty(PropertyIDMap.PID_CATEGORY, value);
        }

        /// <summary>
        /// Removes the category.
        /// </summary>
        public void RemoveCategory()
        {
            Remove1stProperty(PropertyIDMap.PID_CATEGORY);
        }



        /// <summary>
        /// Returns the presentation format (or
        /// <c>null</c>).
        /// </summary>
        /// <return>presentation format value</return>
        public String PresentationFormat
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_PRESFORMAT);
            set => FirstSection.SetProperty(PropertyIDMap.PID_PRESFORMAT, value);
        }

        /// <summary>
        /// Removes the presentation format.
        /// </summary>
        public void RemovePresentationFormat()
        {
            Remove1stProperty(PropertyIDMap.PID_PRESFORMAT);
        }



        /// <summary>
        /// get or set the byte count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a byte count.
        /// </summary>
        /// <return>byteCount value</return>
        public int ByteCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_BYTECOUNT);
            set => Set1stProperty(PropertyIDMap.PID_BYTECOUNT, value);
        }

        /// <summary>
        /// Removes the byte count.
        /// </summary>
        public void RemoveByteCount()
        {
            Remove1stProperty(PropertyIDMap.PID_BYTECOUNT);
        }



        /// <summary>
        /// get or set the line count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a line count.
        /// </summary>
        /// <return>line count value</return>
        public int LineCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_LINECOUNT);
            set => Set1stProperty(PropertyIDMap.PID_LINECOUNT, value);
        }

        /// <summary>
        /// Removes the line count.
        /// </summary>
        public void RemoveLineCount()
        {
            Remove1stProperty(PropertyIDMap.PID_LINECOUNT);
        }



        /// <summary>
        /// get or set the par count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a par count.
        /// </summary>
        /// <return>par count value</return>
        public int ParCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_PARCOUNT);
            set => Set1stProperty(PropertyIDMap.PID_PARCOUNT, value);
        }

        /// <summary>
        /// Removes the par count.
        /// </summary>
        public void RemoveParCount()
        {
            Remove1stProperty(PropertyIDMap.PID_PARCOUNT);
        }



        /// <summary>
        /// get or set the slide count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a slide count.
        /// </summary>
        /// <return>slide count value</return>
        public int SlideCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_SLIDECOUNT);
            set => Set1stProperty(PropertyIDMap.PID_SLIDECOUNT, value);
        }

        /// <summary>
        /// Removes the slide count.
        /// </summary>
        public void RemoveSlideCount()
        {
            Remove1stProperty(PropertyIDMap.PID_SLIDECOUNT);
        }



        /// <summary>
        /// get or set the note count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a note count.
        /// </summary>
        /// <return>note count value</return>
        public int NoteCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_NOTECOUNT);
            set => Set1stProperty(PropertyIDMap.PID_NOTECOUNT, value);
        }

        /// <summary>
        /// Removes the noteCount.
        /// </summary>
        public void RemoveNoteCount()
        {
            Remove1stProperty(PropertyIDMap.PID_NOTECOUNT);
        }



        /// <summary>
        /// get or set the hidden count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a hidden
        /// count.
        /// </summary>
        /// <return>hidden count value</return>
        public int HiddenCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_HIDDENCOUNT);
            set => Set1stProperty(PropertyIDMap.PID_HIDDENCOUNT, value);
        }

        /// <summary>
        /// Removes the hidden count.
        /// </summary>
        public void RemoveHiddenCount()
        {
            Remove1stProperty(PropertyIDMap.PID_HIDDENCOUNT);
        }

        /// <summary>
        /// get or set the mmclip count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a mmclip
        /// count.
        /// </summary>
        /// <return>mmclip count value</return>
        public int MMClipCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_MMCLIPCOUNT);
            set => Set1stProperty(PropertyIDMap.PID_MMCLIPCOUNT, value);
        }


        /// <summary>
        /// Removes the mmclip count.
        /// </summary>
        public void RemoveMMClipCount()
        {
            Remove1stProperty(PropertyIDMap.PID_MMCLIPCOUNT);
        }


        /// <summary>
        /// get or set <c>true</c> when scaling of the thumbnail is
        /// desired, <c>false</c> if cropping is desired.
        /// </summary>
        /// <return>scale value</return>
        public bool Scale
        {
            get => GetPropertyBooleanValue(PropertyIDMap.PID_SCALE);
            set => Set1stProperty(PropertyIDMap.PID_SCALE, value);
        }

        /// <summary>
        /// Removes the scale.
        /// </summary>
        public void RemoveScale()
        {
            Remove1stProperty(PropertyIDMap.PID_SCALE);
        }

        /// <summary>
        /// <para>
        /// Returns the heading pair (or <c>null</c>)
        /// <strong>when this method is implemented. Please note that the
        /// return type is likely to change!</strong>
        /// </para>
        /// </summary>
        /// <return>heading pair value</return>
        public byte[] HeadingPair
        {
            get
            {
                NotYetImplemented("Reading byte arrays ");
                return (byte[]) GetProperty(PropertyIDMap.PID_HEADINGPAIR);
            }
            set
            {
                NotYetImplemented("Writing byte arrays ");
            }
        }

        /// <summary>
        /// Removes the heading pair.
        /// </summary>
        public void RemoveHeadingPair()
        {
            Remove1stProperty(PropertyIDMap.PID_HEADINGPAIR);
        }



        /// <summary>
        /// <para>
        /// Returns the doc parts (or <c>null</c>)
        /// <strong>when this method is implemented. Please note that the
        /// return type is likely to change!</strong>
        /// </para>
        /// </summary>
        /// <return>doc parts value</return>
        public byte[] Docparts
        {
            get
            {
                NotYetImplemented("Reading byte arrays");
                return (byte[]) GetProperty(PropertyIDMap.PID_DOCPARTS);
            }
            set
            {
                NotYetImplemented("Writing byte arrays");
            }
        }

        /// <summary>
        /// Removes the doc parts.
        /// </summary>
        public void RemoveDocparts()
        {
            Remove1stProperty(PropertyIDMap.PID_DOCPARTS);
        }



        /// <summary>
        /// get or set the manager (or <c>null</c>).
        /// </summary>
        /// <return>manager value</return>
        public String Manager
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_MANAGER);
            set => Set1stProperty(PropertyIDMap.PID_MANAGER, value);
        }

        /// <summary>
        /// Removes the manager.
        /// </summary>
        public void RemoveManager()
        {
            Remove1stProperty(PropertyIDMap.PID_MANAGER);
        }



        /// <summary>
        /// Returns the company (or <c>null</c>).
        /// </summary>
        /// <return>company value</return>
        public String Company
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_COMPANY);
            set => Set1stProperty(PropertyIDMap.PID_COMPANY, value);
        }

        /// <summary>
        /// Removes the company.
        /// </summary>
        public void RemoveCompany()
        {
            Remove1stProperty(PropertyIDMap.PID_COMPANY);
        }


        /// <summary>
        /// get <c>true</c> if the custom links are dirty. 
        /// </summary>
        /// <return>links dirty value</return>
        public bool LinksDirty
        {
            get => GetPropertyBooleanValue(PropertyIDMap.PID_LINKSDIRTY);
            set => Set1stProperty(PropertyIDMap.PID_LINKSDIRTY, value);
        }

        /// <summary>
        /// Removes the links dirty.
        /// </summary>
        public void RemoveLinksDirty()
        {
            Remove1stProperty(PropertyIDMap.PID_LINKSDIRTY);
        }


        /// <summary>
        /// <para>
        /// Returns the character count including whitespace, or 0 if the
        /// {@link DocumentSummaryInformation} does not contain this char count.
        /// </para>
        /// <para>
        /// This is the whitespace-including version of {@link SummaryInformation#getCharCount()}
        /// </para>
        /// </summary>
        /// <return>character count or <c>null</c></return>
        public int CharCountWithSpaces
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_CCHWITHSPACES);
            set => Set1stProperty(PropertyIDMap.PID_CCHWITHSPACES, value);
        }

        /// <summary>
        /// Removes the character count
        /// </summary>
        public void RemoveCharCountWithSpaces()
        {
            Remove1stProperty(PropertyIDMap.PID_CCHWITHSPACES);
        }


        /// <summary>
        /// <para>
        /// Get if the User Defined Property Set has been updated outside of the
        /// </para>
        /// <para>
        /// Application.
        /// If it has (true), the hyperlinks should be updated on document load.
        /// </para>
        /// </summary>
        /// <return>if the hyperlinks should be updated on document load</return>
        public bool HyperlinksChanged
        {
            get => GetPropertyBooleanValue(PropertyIDMap.PID_HYPERLINKSCHANGED);
            set => Set1stProperty(PropertyIDMap.PID_HYPERLINKSCHANGED, value);
        }

        /// <summary>
        /// Removes the flag for if the User Defined Property Set has been updated
        /// outside of the Application.
        /// </summary>
        public void RemoveHyperlinksChanged()
        {
            Remove1stProperty(PropertyIDMap.PID_HYPERLINKSCHANGED);
        }


        /// <summary>
        /// <para>
        /// get or set the version of the Application which wrote the
        /// Property Set, stored with the two high order bytes having the major
        /// version number, and the two low order bytes the minor version number.
        /// This will be 0 if no version is Set.
        /// </para>
        /// </summary>
        /// <return>Application version</return>
        public int ApplicationVersion
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_VERSION);
            set => Set1stProperty(PropertyIDMap.PID_VERSION, value);
        }

        /// <summary>
        /// Sets the Application version, which must be a 4 byte int with
        /// the  two high order bytes having the major version number, and the
        /// two low order bytes the minor version number.
        /// </summary>


        /// <summary>
        /// Removes the Application Version
        /// </summary>
        public void RemoveApplicationVersion()
        {
            Remove1stProperty(PropertyIDMap.PID_VERSION);
        }


        /// <summary>
        /// Returns the VBA digital signature for the VBA project
        /// embedded in the document (or <c>null</c>).
        /// </summary>
        /// <return>VBA digital signature</return>
        public byte[] VBADigitalSignature
        {
            get
            {
                Object value = GetProperty(PropertyIDMap.PID_DIGSIG);
                if(value != null && value is byte[] bytes)
                {
                    return bytes;
                }
                return null;
            }
            set
            {
                Set1stProperty(PropertyIDMap.PID_DIGSIG, value);
            }
        }

        /// <summary>
        /// Removes the VBA Digital Signature
        /// </summary>
        public void RemoveVBADigitalSignature()
        {
            Remove1stProperty(PropertyIDMap.PID_DIGSIG);
        }


        /// <summary>
        /// Gets the content type of the file (or <c>null</c>).
        /// </summary>
        /// <return>content type of the file</return>
        public String ContentType
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_CONTENTTYPE);
            set => Set1stProperty(PropertyIDMap.PID_CONTENTTYPE, value);
        }

        /// <summary>
        /// Removes the content type of the file
        /// </summary>
        public void RemoveContentType()
        {
            Remove1stProperty(PropertyIDMap.PID_CONTENTTYPE);
        }


        /// <summary>
        /// Gets the content status of the file (or <c>null</c>).
        /// </summary>
        /// <return>content status of the file</return>
        public String ContentStatus
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_CONTENTSTATUS);
            set => Set1stProperty(PropertyIDMap.PID_CONTENTSTATUS, value);
        }

        /// <summary>
        /// Removes the content status of the file
        /// </summary>
        public void RemoveContentStatus()
        {
            Remove1stProperty(PropertyIDMap.PID_CONTENTSTATUS);
        }


        /// <summary>
        /// Gets the document language, which is normally unset and empty (or <c>null</c>).
        /// </summary>
        /// <return>document language</return>
        public String Language
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_LANGUAGE);
            set => Set1stProperty(PropertyIDMap.PID_LANGUAGE, value);
        }

        /// <summary>
        /// Removes the document language
        /// </summary>
        public void RemoveLanguage()
        {
            Remove1stProperty(PropertyIDMap.PID_LANGUAGE);
        }


        /// <summary>
        /// get or set the document version as a string, which is normally unset and empty
        /// (or <c>null</c>).
        /// </summary>
        /// <return>document verion</return>
        public String DocumentVersion
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_DOCVERSION);
            set => Set1stProperty(PropertyIDMap.PID_DOCVERSION, value);
        }


        /// <summary>
        /// Removes the document version string
        /// </summary>
        public void RemoveDocumentVersion()
        {
            Remove1stProperty(PropertyIDMap.PID_DOCVERSION);
        }


        /// <summary>
        /// get or set the custom properties.
        /// </summary>
        /// <return>custom properties.</return>
        public CustomProperties CustomProperties
        {
            get
            {
                CustomProperties cps = null;
                if(SectionCount >= 2)
                {
                    cps = new CustomProperties();
                    Section section = Sections[1];
                    IDictionary dictionary = section.Dictionary;
                    Property[] properties = section.Properties;
                    int propertyCount = 0;
                    for(int i = 0; i < properties.Length; i++)
                    {
                        Property p = properties[i];
                        long id = p.ID;
                        if(id != 0 && id != 1)
                        {
                            propertyCount++;
                            CustomProperty cp = new CustomProperty(p,
                                (string)dictionary[id]);
                            cps.Put(cp.Name, cp);
                        }
                    }
                    if(cps.Count != propertyCount)
                    {
                        cps.IsPure = (false);
                    }
                }
                return cps;
            }
            set
            {
                ensureSection2();
                Section section = Sections[1];
                Dictionary<long, String> dictionary = value.GetDictionary();
                section.clear();

                /* Set the codepage. If both custom properties and section have a
                 * codepage, the codepage from the custom properties wins, else take the
                 * one that is defined. If none is defined, take Unicode. */
                int cpCodepage = value.GetCodepage();
                if(cpCodepage < 0)
                {
                    cpCodepage = section.Codepage;
                }
                if(cpCodepage < 0)
                {
                    cpCodepage = CodePageUtil.CP_UNICODE;
                }
                value.SetCodepage(cpCodepage);
                section.SetCodepage(cpCodepage);
                section.SetDictionary(dictionary);
                //i = section.Size;
                foreach(CustomProperty p in value.Values)
                {
                    section.SetProperty(p);
                }
            }
        }

        /// <summary>
        /// Creates section 2 if it is not already present.
        /// </summary>
        private void ensureSection2()
        {
            if(SectionCount < 2)
            {
                Section s2 = new MutableSection();
                s2.SetFormatID(SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID[1]);
                AddSection(s2);
            }
        }

        /// <summary>
        /// Removes the custom properties.
        /// </summary>
        public void RemoveCustomProperties()
        {
            if(SectionCount < 2)
            {
                throw new HPSFRuntimeException("Illegal internal format of Document SummaryInformation stream: second section is missing.");
            }

            List<Section> l = new List<Section>(Sections);
            ClearSections();
            int idx = 0;
            foreach(Section s in l)
            {
                if(idx++ != 1)
                {
                    AddSection(s);
                }
            }
        }

        /// <summary>
        /// Throws an {@link UnsupportedOperationException} with a message text
        /// telling which functionality is not yet implemented.
        /// </summary>
        /// <param name="msg">text telling was leaves to be implemented, e.g.
        /// "Reading byte arrays".
        /// </param>
        private void NotYetImplemented(String msg)
        {
            throw new NotImplementedException(msg + " is not yet implemented.");
        }
    }
}

