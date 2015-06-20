/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

namespace NPOI.HPSF
{
    using System;
    using System.Collections;
    using NPOI.HPSF.Wellknown;
    using NPOI.Util;

    /// <summary>
    /// Convenience class representing a DocumentSummary Information stream in a
    /// Microsoft Office document.
    /// @author Rainer Klute 
    /// klute@rainer-klute.de
    /// @author Drew Varner (Drew.Varner cloSeto sc.edu)
    /// @author robert_flaherty@hyperion.com
    /// @since 2002-02-09
    /// </summary>
    [Serializable]
    public class DocumentSummaryInformation : SpecialPropertySet
    {

        /**
         * The document name a document summary information stream
         * usually has in a POIFS filesystem.
         */
        public const string DEFAULT_STREAM_NAME = "\x0005DocumentSummaryInformation";

        public override PropertyIDMap PropertySetIDMap
        {
            get
            {
                return PropertyIDMap.DocumentSummaryInformationProperties;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentSummaryInformation"/> class.
        /// </summary>
        /// <param name="ps">A property Set which should be Created from a
        /// document summary information stream.</param>
        public DocumentSummaryInformation(PropertySet ps): base(ps)
        {
            if (!IsDocumentSummaryInformation)
                throw new UnexpectedPropertySetTypeException
                    ("Not a " + GetType().Name);
        }



        /// <summary>
        /// Gets or sets the category.
        /// </summary>
        /// <value>The category value</value>
        public String Category
        {
            get
            {
                return GetPropertyStringValue(PropertyIDMap.PID_CATEGORY);
            }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_CATEGORY, value);
            }
        }

        /// <summary>
        /// Removes the category.
        /// </summary>
        public void RemoveCategory()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_CATEGORY);
        }



        /// <summary>
        /// Gets or sets the presentation format (or null).
        /// </summary>
        /// <value>The presentation format value</value>
        public String PresentationFormat
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_PRESFORMAT); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_PRESFORMAT, value);
            }
        }

        /// <summary>
        /// Removes the presentation format.
        /// </summary>
        public void RemovePresentationFormat()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_PRESFORMAT);
        }



        /// <summary>
        /// Gets or sets the byte count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a byte count.
        /// </summary>
        /// <value>The byteCount value</value>
        public int ByteCount
        {
            get
            {
                return GetPropertyIntValue(PropertyIDMap.PID_BYTECOUNT);
            }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_BYTECOUNT, value);
            }
        }

        /// <summary>
        /// Removes the byte count.
        /// </summary>
        public void RemoveByteCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_BYTECOUNT);
        }



        /// <summary>
        /// Gets or sets the line count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a line count.
        /// </summary>
        /// <value>The line count value.</value>
        public int LineCount
        {
            get
            {
                return GetPropertyIntValue(PropertyIDMap.PID_LINECOUNT);
            }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_LINECOUNT, value);
            }
        }


        /// <summary>
        /// Removes the line count.
        /// </summary>
        public void RemoveLineCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_LINECOUNT);
        }



        /// <summary>
        /// Gets or sets the par count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a par count.
        /// </summary>
        /// <value>The par count value</value>
        public int ParCount
        {
            get { return GetPropertyIntValue(PropertyIDMap.PID_PARCOUNT); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_PARCOUNT, value);
            }
        }


        /// <summary>
        /// Removes the par count.
        /// </summary>
        public void RemoveParCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_PARCOUNT);
        }



        /// <summary>
        /// Gets or sets the slide count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a slide count.
        /// </summary>
        /// <value>The slide count value</value>
        public int SlideCount
        {
            get { return GetPropertyIntValue(PropertyIDMap.PID_SLIDECOUNT); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_SLIDECOUNT, value);
            }
        }

        /// <summary>
        /// Removes the slide count.
        /// </summary>
        public void RemoveSlideCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_SLIDECOUNT);
        }



        /// <summary>
        /// Gets or sets the note count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a note count
        /// </summary>
        /// <value>The note count value</value>
        public int NoteCount
        {
            get { return GetPropertyIntValue(PropertyIDMap.PID_NOTECOUNT); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_NOTECOUNT, value);
            }
        }

        /// <summary>
        /// Removes the note count.
        /// </summary>
        public void RemoveNoteCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_NOTECOUNT);
        }



        /// <summary>
        /// Gets or sets the hidden count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a hidden
        /// count.
        /// </summary>
        /// <value>The hidden count value.</value>
        public int HiddenCount
        {
            get { return GetPropertyIntValue(PropertyIDMap.PID_HIDDENCOUNT); }
            set
            {
                MutableSection s = (MutableSection)Sections[0];
                s.SetProperty(PropertyIDMap.PID_HIDDENCOUNT, value);
            }
        }

        /// <summary>
        /// Removes the hidden count.
        /// </summary>
        public void RemoveHiddenCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_HIDDENCOUNT);
        }



        /// <summary>
        /// Returns the mmclip count or 0 if the {@link
        /// DocumentSummaryInformation} does not contain a mmclip
        /// count.
        /// </summary>
        /// <value>The mmclip count value.</value>
        public int MMClipCount
        {
            get { return GetPropertyIntValue(PropertyIDMap.PID_MMCLIPCOUNT); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_MMCLIPCOUNT, value);
            }
        }

        /// <summary>
        /// Removes the MMClip count.
        /// </summary>
        public void RemoveMMClipCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_MMCLIPCOUNT);
        }



        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="DocumentSummaryInformation"/> is scale.
        /// </summary>
        /// <value><c>true</c> if cropping is desired; otherwise, <c>false</c>.</value>
        public bool Scale
        {
            get { return GetPropertyBooleanValue(PropertyIDMap.PID_SCALE); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_SCALE, value);
            }
        }

        /// <summary>
        /// Removes the scale.
        /// </summary>
        public void RemoveScale()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_SCALE);
        }



        /// <summary>
        /// Gets or sets the heading pair (or null)
        /// </summary>
        /// <value>The heading pair value.</value>
        public byte[] HeadingPair
        {
            get
            {
                return (byte[])GetProperty(PropertyIDMap.PID_HEADINGPAIR);
            }
            set {
                throw new NotImplementedException("Writing byte arrays ");
            }
        }

        /// <summary>
        /// Removes the heading pair.
        /// </summary>
        public void RemoveHeadingPair()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_HEADINGPAIR);
        }



        /// <summary>
        /// Gets or sets the doc parts.
        /// </summary>
        /// <value>The doc parts value</value>
        public byte[] Docparts
        {
            get
            {
                return (byte[])GetProperty(PropertyIDMap.PID_DOCPARTS);
            }
            set 
            {
                throw new NotImplementedException("Writing byte arrays");
            }
        }

        /// <summary>
        /// Removes the doc parts.
        /// </summary>
        public void RemoveDocparts()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_DOCPARTS);
        }



        /// <summary>
        /// Gets or sets the manager (or <c>null</c>).
        /// </summary>
        /// <value>The manager value</value>
        public String Manager
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_MANAGER); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_MANAGER, value);
            }
        }

        /// <summary>
        /// Removes the manager.
        /// </summary>
        public void RemoveManager()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_MANAGER);
        }



        /// <summary>
        /// Gets or sets the company (or <c>null</c>).
        /// </summary>
        /// <value>The company value</value>
        public String Company
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_COMPANY); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_COMPANY, value);
            }
        }

        /// <summary>
        /// Removes the company.
        /// </summary>
        public void RemoveCompany()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_COMPANY);
        }



        /// <summary>
        /// Gets or sets a value indicating whether [links dirty].
        /// </summary>
        /// <value><c>true</c> if the custom links are dirty.; otherwise, <c>false</c>.</value>
        public bool LinksDirty
        {
            get { return GetPropertyBooleanValue(PropertyIDMap.PID_LINKSDIRTY); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_LINKSDIRTY, value);
            }
        }

        /// <summary>
        /// Removes the links dirty.
        /// </summary>
        public void RemoveLinksDirty()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_LINKSDIRTY);
        }



        /// <summary>
        /// Gets or sets the custom properties.
        /// </summary>
        /// <value>The custom properties.</value>
        public CustomProperties CustomProperties
        {
            get
            {
                CustomProperties cps = null;
                if (SectionCount >= 2)
                {
                    cps = new CustomProperties();
                    Section section = (Section)Sections[1];
                    IDictionary dictionary = section.Dictionary;
                    Property[] properties = section.Properties;
                    int propertyCount = 0;
                    for (int i = 0; i < properties.Length; i++)
                    {
                        Property p = properties[i];
                        long id = p.ID;
                        if (id != 0 && id != 1)
                        {
                            propertyCount++;
                            CustomProperty cp = new CustomProperty(p,
                                    (string)dictionary[id]);
                            cps.Put(cp.Name, cp);
                        }
                    }
                    if (cps.Count != propertyCount)
                        cps.IsPure=false;
                }
                return cps;
            }
            set 
            {
                EnsureSection2();
                MutableSection section = (MutableSection)Sections[1];
                IDictionary dictionary = value.Dictionary;
                section.Clear();

                /* Set the codepage. If both custom properties and section have a
                 * codepage, the codepage from the custom properties wins, else take the
                 * one that is defined. If none is defined, take Unicode. */
                int cpCodepage = value.Codepage;
                if (cpCodepage < 0)
                    cpCodepage = section.Codepage;
                if (cpCodepage < 0)
                    cpCodepage = CodePageUtil.CP_UNICODE;
                value.Codepage=cpCodepage;
                section.Codepage=cpCodepage; //add codepage propertyset
                section.Dictionary=dictionary; //generate dictionary propertyset
                //generate MutableSections
                for (IEnumerator i = value.Values.GetEnumerator(); i.MoveNext(); )
                {
                    Property p = (Property)i.Current;
                    section.SetProperty(p);
                }
            }
        }


        /// <summary>
        /// Creates section 2 if it is not alReady present.
        /// </summary>
        private void EnsureSection2()
        {
            if (SectionCount < 2)
            {
                MutableSection s2 = new MutableSection();
                s2.SetFormatID(SectionIDMap.DOCUMENT_SUMMARY_INFORMATION_ID2);
                AddSection(s2);
            }
        }



        /// <summary>
        /// Removes the custom properties.
        /// </summary>
        public void RemoveCustomProperties()
        {
            if (SectionCount >= 2)
                Sections.RemoveAt(1);
            else
                throw new HPSFRuntimeException("Illegal internal format of Document SummaryInformation stream: second section is missing.");
        }

    }
}