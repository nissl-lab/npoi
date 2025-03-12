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

namespace NPOI.HPSF

{
    using NPOI.HPSF.Wellknown;
    using System;
    using System.Xml.Linq;

    /// <summary>
    /// Convenience class representing a Summary Information stream in a
    /// Microsoft Office document.
    /// </summary>
    /// @see DocumentSummaryInformation
    public class SummaryInformation : SpecialPropertySet
    {

        /// <summary>
        /// The document name a summary information stream usually has in a POIFS filesystem.
        /// </summary>
        public static String DEFAULT_STREAM_NAME = "\x0005SummaryInformation";

        public override PropertyIDMap PropertySetIDMap => PropertyIDMap.SummaryInformationProperties;

        /// <summary>
        /// Creates an empty {@link SummaryInformation}.
        /// </summary>
        public SummaryInformation()
        {
            FirstSection.SetFormatID(SectionIDMap.SUMMARY_INFORMATION_ID);
        }

        /// <summary>
        /// Creates a {@link SummaryInformation} from a given {@link
        /// PropertySet}.
        /// </summary>
        /// <param name="ps">A property Set which should be created from a summary
        /// information stream.
        /// </param>
        /// <exception name="UnexpectedPropertySetTypeException">if <c>ps</c> does not
        /// contain a summary information stream.
        /// </exception>
        public SummaryInformation(PropertySet ps) : base(ps)
        {
            if(!IsSummaryInformation)
            {
                throw new UnexpectedPropertySetTypeException("Not a " + GetType().Name);
            }
        }



        /// <summary>
        /// get or set the title.
        /// </summary>
        /// <return>title or <c>null</c></return>
        public String Title
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_TITLE);
            set => Set1stProperty(PropertyIDMap.PID_TITLE, value);
        }


        /// <summary>
        /// Removes the title.
        /// </summary>
        public void RemoveTitle()
        {
            Remove1stProperty(PropertyIDMap.PID_TITLE);
        }



        /// <summary>
        /// get or set the subject (or <c>null</c>).
        /// </summary>
        /// <return>subject or <c>null</c></return>
        public String Subject
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_SUBJECT);
            set => Set1stProperty(PropertyIDMap.PID_SUBJECT, value);
        }

        /// <summary>
        /// Removes the subject.
        /// </summary>
        public void RemoveSubject()
        {
            Remove1stProperty(PropertyIDMap.PID_SUBJECT);
        }



        /// <summary>
        /// get or set the author (or <c>null</c>).
        /// </summary>
        /// <return>author or <c>null</c></return>
        public String Author
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_AUTHOR);
            set => Set1stProperty(PropertyIDMap.PID_AUTHOR, value);
        }


        /// <summary>
        /// Removes the author.
        /// </summary>
        public void RemoveAuthor()
        {
            Remove1stProperty(PropertyIDMap.PID_AUTHOR);
        }



        /// <summary>
        /// get or set the keywords (or <c>null</c>).
        /// </summary>
        /// <return>keywords or <c>null</c></return>
        public String Keywords
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_KEYWORDS);
            set => Set1stProperty(PropertyIDMap.PID_KEYWORDS, value);
        }


        /// <summary>
        /// Removes the keywords.
        /// </summary>
        public void RemoveKeywords()
        {
            Remove1stProperty(PropertyIDMap.PID_KEYWORDS);
        }



        /// <summary>
        /// get or set the comments (or <c>null</c>).
        /// </summary>
        /// <return>comments or <c>null</c></return>
        public String Comments
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_COMMENTS);
            set => Set1stProperty(PropertyIDMap.PID_COMMENTS, value);
        }

        /// <summary>
        /// Removes the comments.
        /// </summary>
        public void RemoveComments()
        {
            Remove1stProperty(PropertyIDMap.PID_COMMENTS);
        }


        /// <summary>
        /// get or set the template (or <c>null</c>).
        /// </summary>
        /// <return>template or <c>null</c></return>
        public String Template
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_TEMPLATE);
            set => Set1stProperty(PropertyIDMap.PID_TEMPLATE, value);
        }

        /// <summary>
        /// Removes the template.
        /// </summary>
        public void RemoveTemplate()
        {
            Remove1stProperty(PropertyIDMap.PID_TEMPLATE);
        }



        /// <summary>
        /// get or set the last author (or <c>null</c>).
        /// </summary>
        /// <return>last author or <c>null</c></return>
        public String LastAuthor
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_LASTAUTHOR);
            set => Set1stProperty(PropertyIDMap.PID_LASTAUTHOR, value);
        }

        /// <summary>
        /// Removes the last author.
        /// </summary>
        public void RemoveLastAuthor()
        {
            Remove1stProperty(PropertyIDMap.PID_LASTAUTHOR);
        }



        /// <summary>
        /// get or set the revision number (or <c>null</c>).
        /// </summary>
        /// <return>revision number or <c>null</c></return>
        public String RevNumber
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_REVNUMBER);
            set => Set1stProperty(PropertyIDMap.PID_REVNUMBER, value);
        }

        /// <summary>
        /// Removes the revision number.
        /// </summary>
        public void RemoveRevNumber()
        {
            Remove1stProperty(PropertyIDMap.PID_REVNUMBER);
        }



        /// <summary>
        /// get or set the total time spent in editing the document (or
        /// <c>0</c>).
        /// </summary>
        /// <return>total time spent in editing the document or 0 if the {@link
        /// SummaryInformation} does not contain this information.
        /// </return>
        public long EditTime
        {
            get
            {
                DateTime? d = (DateTime?)GetProperty(PropertyIDMap.PID_EDITTIME);

                return d.HasValue ? Util.DateToFileTime(d.Value) : 0;
            }
            set
            {
                DateTime d = Util.FiletimeToDate(value);
                FirstSection.SetProperty(PropertyIDMap.PID_EDITTIME, Variant.VT_FILETIME, d);
            }
        }

        /// <summary>
        /// Remove the total time spent in editing the document.
        /// </summary>
        public void RemoveEditTime()
        {
            Remove1stProperty(PropertyIDMap.PID_EDITTIME);
        }



        /// <summary>
        /// get or set the last printed time (or <c>null</c>).
        /// </summary>
        /// <return>last printed time or <c>null</c></return>
        public DateTime? LastPrinted
        {
            get => (DateTime?) GetProperty(PropertyIDMap.PID_LASTPRINTED);
            set => FirstSection.SetProperty(PropertyIDMap.PID_LASTPRINTED, Variant.VT_FILETIME, value);
        }

        /// <summary>
        /// Removes the lastPrinted.
        /// </summary>
        public void RemoveLastPrinted()
        {
            Remove1stProperty(PropertyIDMap.PID_LASTPRINTED);
        }



        /// <summary>
        /// get or set the creation time (or <c>null</c>).
        /// </summary>
        /// <return>creation time or <c>null</c></return>
        public DateTime? CreateDateTime
        {
            get => (DateTime?) GetProperty(PropertyIDMap.PID_CREATE_DTM);
            set => FirstSection.SetProperty(PropertyIDMap.PID_CREATE_DTM, Variant.VT_FILETIME, value);
        }


        /// <summary>
        /// Removes the creation time.
        /// </summary>
        public void RemoveCreateDateTime()
        {
            Remove1stProperty(PropertyIDMap.PID_CREATE_DTM);
        }



        /// <summary>
        /// get or set the last save time (or <c>null</c>).
        /// </summary>
        /// <return>last save time or <c>null</c></return>
        public DateTime? LastSaveDateTime
        {
            get => (DateTime?) GetProperty(PropertyIDMap.PID_LASTSAVE_DTM);
            set => FirstSection.SetProperty(PropertyIDMap.PID_LASTSAVE_DTM, Variant.VT_FILETIME, value);
        }


        /// <summary>
        /// Remove the total time spent in editing the document.
        /// </summary>
        public void RemoveLastSaveDateTime()
        {
            Remove1stProperty(PropertyIDMap.PID_LASTSAVE_DTM);
        }



        /// <summary>
        /// get or set the page count or 0 if the {@link SummaryInformation} does
        /// not contain a page count.
        /// </summary>
        /// <return>page count or 0 if the {@link SummaryInformation} does not
        /// contain a page count.
        /// </return>
        public int PageCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_PAGECOUNT);
            set => Set1stProperty(PropertyIDMap.PID_PAGECOUNT, value);
        }

        /// <summary>
        /// Removes the page count.
        /// </summary>
        public void RemovePageCount()
        {
            Remove1stProperty(PropertyIDMap.PID_PAGECOUNT);
        }



        /// <summary>
        /// get or set the word count or 0 if the {@link SummaryInformation} does
        /// not contain a word count.
        /// </summary>
        /// <return>word count or <c>null</c></return>
        public int WordCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_WORDCOUNT);
            set => Set1stProperty(PropertyIDMap.PID_WORDCOUNT, value);
        }


        /// <summary>
        /// Removes the word count.
        /// </summary>
        public void RemoveWordCount()
        {
            Remove1stProperty(PropertyIDMap.PID_WORDCOUNT);
        }



        /// <summary>
        /// get or set the character count or 0 if the {@link SummaryInformation}
        /// does not contain a char count.
        /// </summary>
        /// <return>character count or <c>null</c></return>
        public int CharCount
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_CHARCOUNT);
            set => Set1stProperty(PropertyIDMap.PID_CHARCOUNT, value);
        }

        /// <summary>
        /// Removes the character count.
        /// </summary>
        public void RemoveCharCount()
        {
            Remove1stProperty(PropertyIDMap.PID_CHARCOUNT);
        }



        /// <summary>
        /// <para>
        /// get or set the thumbnail (or <c>null</c>) <strong>when this
        /// method is implemented. Please note that the return type is likely to
        /// change!</strong>
        /// </para>
        /// <para>
        /// To process this data, you may wish to make use of the
        /// {@link Thumbnail} class. The raw data is generally
        /// an image in WMF or Clipboard (BMP?) format
        /// </para>
        /// </summary>
        /// <return>thumbnail or <c>null</c></return>
        public byte[] Thumbnail
        {
            get => (byte[]) GetProperty(PropertyIDMap.PID_THUMBNAIL);
            set => FirstSection.SetProperty(PropertyIDMap.PID_THUMBNAIL, /* FIXME: */ Variant.VT_LPSTR, value);
        }

        /// <summary>
        /// get or set the thumbnail (or <c>null</c>), processed
        /// as an object which is (largely) able to unpack the thumbnail
        /// image data.
        /// </summary>
        /// <return>thumbnail or <c>null</c></return>
        public Thumbnail ThumbnailThumbnail
        {
            get
            {
                byte[] data = Thumbnail;
                if(data == null)
                    return null;
                return new Thumbnail(data);
            }

        }


        /// <summary>
        /// Removes the thumbnail.
        /// </summary>
        public void RemoveThumbnail()
        {
            Remove1stProperty(PropertyIDMap.PID_THUMBNAIL);
        }



        /// <summary>
        /// get or set the application name (or <c>null</c>).
        /// </summary>
        /// <return>application name or <c>null</c></return>
        public String ApplicationName
        {
            get => GetPropertyStringValue(PropertyIDMap.PID_APPNAME);
            set => Set1stProperty(PropertyIDMap.PID_APPNAME, value);
        }


        /// <summary>
        /// Removes the application name.
        /// </summary>
        public void RemoveApplicationName()
        {
            Remove1stProperty(PropertyIDMap.PID_APPNAME);
        }



        /// <summary>
        /// <para>
        /// get or set a security code which is one of the following values:
        /// </para>
        /// <para>
        /// <ul>
        /// </para>
        /// <para>
        /// <li>0 if the {@link SummaryInformation} does not contain a
        /// security field or if there is no security on the document. Use
        /// {@link PropertySet#wasNull()} to distinguish between the two
        /// cases!
        /// </para>
        /// <para>
        /// <li>1 if the document is password protected
        /// </para>
        /// <para>
        /// <li>2 if the document is read-only recommended
        /// </para>
        /// <para>
        /// <li>4 if the document is read-only enforced
        /// </para>
        /// <para>
        /// <li>8 if the document is locked for annotations
        /// </para>
        /// <para>
        /// </ul>
        /// </para>
        /// </summary>
        /// <return>security code or <c>null</c></return>
        public int Security
        {
            get => GetPropertyIntValue(PropertyIDMap.PID_SECURITY);
            set => Set1stProperty(PropertyIDMap.PID_SECURITY, value);
        }

        /// <summary>
        /// Removes the security code.
        /// </summary>
        public void RemoveSecurity()
        {
            Remove1stProperty(PropertyIDMap.PID_SECURITY);
        }

    }
}

