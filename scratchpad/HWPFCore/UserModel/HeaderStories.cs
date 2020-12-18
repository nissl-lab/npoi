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

using NPOI.HWPF.UserModel;
using NPOI.HWPF.Model;
using System;
namespace NPOI.HWPF.UserModel
{

    /**
     * A HeaderStory is a Header, a Footer, or footnote/endnote
     *  separator.
     * All the Header Stories get stored in the same Range in the
     *  document, and this handles Getting out all the individual
     *  parts.
     *
     * WARNING - you shouldn't change the headers or footers,
     *  as offSets are not yet updated!
     */
    public class HeaderStories
    {
        private Range headerStories;
        private PlexOfCps plcfHdd;

        private bool stripFields = false;

        public HeaderStories(HWPFDocument doc)
        {
            this.headerStories = doc.GetHeaderStoryRange();
            FileInformationBlock fib = doc.GetFileInformationBlock();

            // If there's no PlcfHdd, nothing to do
            if (fib.GetCcpHdd() == 0)
            {
                return;
            }
            if (fib.GetPlcfHddSize() == 0)
            {
                return;
            }

            // Handle the PlcfHdd
            plcfHdd = new PlexOfCps(
                    doc.GetTableStream(), fib.GetPlcfHddOffset(),
                    fib.GetPlcfHddSize(), 0
            );
        }

        public String FootnoteSeparator
        {
            get{
                return GetAt(0);
            }
        }
        public String FootnoteContSeparator
        {
            get{
                return GetAt(1);
            }
        }
        public String FootnoteContNote
        {
            get{
            return GetAt(2);
            }
        }
        public String EndnoteSeparator
        {
            get{
            return GetAt(3);
            }
        }
        public String EndnoteContSeparator
        {
            get{
            return GetAt(4);
            }
        }
        public String EndnoteContNote
        {
            get
            {
                return GetAt(5);
            }
        }


        public String EvenHeader
        {
            get
            {
                return GetAt(6 + 0);
            }
        }
        public String OddHeader
        {
            get
            {
                return GetAt(6 + 1);
            }
        }
        public String FirstHeader
        {
            get
            {
                return GetAt(6 + 4);
            }
        }
        /**
         * Returns the correct, defined header for the given
         *  one based page
         * @param pageNumber The one based page number
         */
        public String GetHeader(int pageNumber)
        {
            // First page header is optional, only return
            //  if it's set
            if (pageNumber == 1)
            {
                if (FirstHeader.Length > 0)
                {
                    return FirstHeader;
                }
            }
            // Even page header is optional, only return
            //  if it's set
            if (pageNumber % 2 == 0)
            {
                if (EvenHeader.Length > 0)
                {
                    return EvenHeader;
                }
            }
            // Odd is the default
            return OddHeader;
        }


        public String EvenFooter
        {
            get
            {
                return GetAt(6 + 2);
            }
        }
        public String OddFooter
        {
            get
            {
                return GetAt(6 + 3);
            }
        }
        public String FirstFooter
        {
            get
            {
                return GetAt(6 + 5);
            }
        }
        /**
         * Returns the correct, defined footer for the given
         *  one based page
         * @param pageNumber The one based page number
         */
        public String GetFooter(int pageNumber)
        {
            // First page footer is optional, only return
            //  if it's set
            if (pageNumber == 1)
            {
                if (FirstFooter.Length > 0)
                {
                    return FirstFooter;
                }
            }
            // Even page footer is optional, only return
            //  if it's set
            if (pageNumber % 2 == 0)
            {
                if (EvenFooter.Length > 0)
                {
                    return EvenFooter;
                }
            }
            // Odd is the default
            return OddFooter;
        }


        /**
         * Get the string that's pointed to by the
         *  given plcfHdd index
         */
        private String GetAt(int plcfHddIndex)
        {
            if (plcfHdd == null) return null;

            GenericPropertyNode prop = plcfHdd.GetProperty(plcfHddIndex);
            if (prop.Start == prop.End)
            {
                // Empty story
                return "";
            }
            if (prop.End < prop.Start)
            {
                // Broken properties?
                return "";
            }

            // Ensure we're Getting a sensible length
            String rawText = headerStories.Text;
            int start = Math.Min(prop.Start, rawText.Length);
            int end = Math.Min(prop.End, rawText.Length);

            // Grab the contents
            String text = rawText.Substring(start, end-start);

            // Strip off fields and macros if requested
            if (stripFields)
            {
                return Range.StripFields(text);
            }
            // If you create a header/footer, then remove it again, word
            //  will leave \r\r. Turn these back into an empty string,
            //  which is more what you'd expect
            if (text.Equals("\r\r"))
            {
                return "";
            }

            return text;
        }

        public Range GetRange()
        {
            return headerStories;
        }
        internal PlexOfCps GetPlcfHdd()
        {
            return plcfHdd;
        }

        /**
         * Are fields currently being stripped from
         *  the text that this {@link HeaderStories} returns?
         *  Default is false, but can be Changed
         */
        public bool AreFieldsStripped
        {
            get
            {
                return stripFields;
            }
            set 
            {
                this.stripFields = value;
            }
        }
    }
}
