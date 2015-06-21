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
    using NPOI.HPSF.Wellknown;


    /// <summary>
    /// Convenience class representing a Summary Information stream in a
    /// Microsoft Office document.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @see DocumentSummaryInformation
    /// @since 2002-02-09
    /// </summary>
    [Serializable]
    public class SummaryInformation : SpecialPropertySet
    {

        /**
         * The document name a summary information stream usually has in a POIFS
         * filesystem.
         */
        public const String DEFAULT_STREAM_NAME = "\x0005SummaryInformation";


        public override PropertyIDMap PropertySetIDMap
        {
            get
            {
                return PropertyIDMap.SummaryInformationProperties;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SummaryInformation"/> class.
        /// </summary>
        /// <param name="ps">A property Set which should be Created from a summary
        /// information stream.</param>
        public SummaryInformation(PropertySet ps): base(ps)
        {

            if (!IsSummaryInformation)
                throw new UnexpectedPropertySetTypeException("Not a "
                        + GetType().Name);
        }



        /// <summary>
        /// Gets or sets the title.
        /// </summary>
        /// <value>The title.</value>
        public String Title
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_TITLE); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_TITLE, value);
            }
        }

        /// <summary>
        /// Removes the title.
        /// </summary>
        public void RemoveTitle()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_TITLE);
        }


        /// <summary>
        /// Gets or sets the subject.
        /// </summary>
        /// <value>The subject.</value>
        public String Subject
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_SUBJECT); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_SUBJECT, value);
            }
        }

        /// <summary>
        /// Removes the subject.
        /// </summary>
        public void RemoveSubject()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_SUBJECT);
        }


        /// <summary>
        /// Gets or sets the author.
        /// </summary>
        /// <value>The author.</value>
        public String Author
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_AUTHOR); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_AUTHOR, value);
            }
        }

        /// <summary>
        /// Removes the author.
        /// </summary>
        public void RemoveAuthor()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_AUTHOR);
        }


        /// <summary>
        /// Gets or sets the keywords.
        /// </summary>
        /// <value>The keywords.</value>
        public String Keywords
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_KEYWORDS); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_KEYWORDS, value);
            }
        }

        /// <summary>
        /// Removes the keywords.
        /// </summary>
        public void RemoveKeywords()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_KEYWORDS);
        }



        /// <summary>
        /// Gets or sets the comments.
        /// </summary>
        /// <value>The comments.</value>
        public String Comments
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_COMMENTS); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_COMMENTS, value);
            }
        }

        /// <summary>
        /// Removes the comments.
        /// </summary>
        public void RemoveComments()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_COMMENTS);
        }

        /// <summary>
        /// Gets or sets the template.
        /// </summary>
        /// <value>The template.</value>
        public String Template
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_TEMPLATE); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_TEMPLATE, value);
            }
        }

        /// <summary>
        /// Removes the template.
        /// </summary>
        public void RemoveTemplate()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_TEMPLATE);
        }

        /// <summary>
        /// Gets or sets the last author.
        /// </summary>
        /// <value>The last author.</value>
        public String LastAuthor
        {
            get{return GetPropertyStringValue(PropertyIDMap.PID_LASTAUTHOR);}
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_LASTAUTHOR, value);
            }
        }

        /// <summary>
        /// Removes the last author.
        /// </summary>
        public void RemoveLastAuthor()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_LASTAUTHOR);
        }


        /// <summary>
        /// Gets or sets the rev number.
        /// </summary>
        /// <value>The rev number.</value>
        public String RevNumber
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_REVNUMBER); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_REVNUMBER, value);
            }
        }

        /// <summary>
        /// Removes the rev number.
        /// </summary>
        public void RemoveRevNumber()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_REVNUMBER);
        }



        /// <summary>
        /// Returns the Total time spent in editing the document (or 0).
        /// </summary>
        /// <value>The Total time spent in editing the document or 0 if the {@link
        /// SummaryInformation} does not contain this information.</value>
        public long EditTime
        {
            get
            {
                if (GetProperty(PropertyIDMap.PID_EDITTIME) == null)
                    return 0;
                else
                {
                    DateTime d = (DateTime)GetProperty(PropertyIDMap.PID_EDITTIME);
                    return Util.DateToFileTime(d);
                }
            }
            set
            {
                DateTime d = Util.FiletimeToDate(value);
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_EDITTIME, Variant.VT_FILETIME, d);
            }
        }


        /// <summary>
        /// Removes the edit time.
        /// </summary>
        public void RemoveEditTime()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_EDITTIME);
        }



        /// <summary>
        /// Gets or sets the last printed time
        /// </summary>
        /// <value>The last printed time</value>
        /// Returns the last printed time (or <c>null</c>).
        public DateTime? LastPrinted
        {
            get { 
                return (DateTime?)GetProperty(PropertyIDMap.PID_LASTPRINTED); 
            }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_LASTPRINTED, Variant.VT_FILETIME,
                        value);
            }
        }

        /// <summary>
        /// Removes the last printed.
        /// </summary>
        public void RemoveLastPrinted()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_LASTPRINTED);
        }

        /// <summary>
        /// Gets or sets the create date time.
        /// </summary>
        /// <value>The create date time.</value>
        public DateTime? CreateDateTime
        {
            get { return (DateTime?)GetProperty(PropertyIDMap.PID_Create_DTM); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_Create_DTM, Variant.VT_FILETIME,
                        value);
            }
        }

        /// <summary>
        /// Removes the create date time.
        /// </summary>
        public void RemoveCreateDateTime()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_Create_DTM);
        }


        /// <summary>
        /// Gets or sets the last save date time.
        /// </summary>
        /// <value>The last save date time.</value>
        public DateTime? LastSaveDateTime
        {
            get { return (DateTime?)GetProperty(PropertyIDMap.PID_LASTSAVE_DTM); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_LASTSAVE_DTM,
                                Variant.VT_FILETIME, value);
            }
        }

        /// <summary>
        /// Removes the last save date time.
        /// </summary>
        public void RemoveLastSaveDateTime()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_LASTSAVE_DTM);
        }



        /// <summary>
        /// Gets or sets the page count or 0 if the {@link SummaryInformation} does
        /// not contain a page count.
        /// </summary>
        /// <value>The page count or 0 if the {@link SummaryInformation} does not
        /// contain a page count.</value>
        public int PageCount
        {
            get { return GetPropertyIntValue(PropertyIDMap.PID_PAGECOUNT); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_PAGECOUNT, value);
            }
        }


        /// <summary>
        /// Removes the page count.
        /// </summary>
        public void RemovePageCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_PAGECOUNT);
        }



        /// <summary>
        /// Gets or sets the word count or 0 if the {@link SummaryInformation} does
        /// not contain a word count.
        /// </summary>
        /// <value>The word count.</value>
        public int WordCount
        {
            get { return GetPropertyIntValue(PropertyIDMap.PID_WORDCOUNT); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_WORDCOUNT, value);
            }
        }

        /// <summary>
        /// Removes the word count.
        /// </summary>
        public void RemoveWordCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_WORDCOUNT);
        }



        /// <summary>
        /// Gets or sets the character count or 0 if the {@link SummaryInformation}
        /// does not contain a char count.
        /// </summary>
        /// <value>The character count.</value>
        public int CharCount
        {
            get{return GetPropertyIntValue(PropertyIDMap.PID_CHARCOUNT);}
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_CHARCOUNT, value);
            }
        }

        /// <summary>
        /// Removes the char count.
        /// </summary>
        public void RemoveCharCount()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_CHARCOUNT);
        }


        /// <summary>
        /// Gets or sets the thumbnail (or <c>null</c>) <strong>when this
        /// method is implemented. Please note that the return type is likely To
        /// Change!</strong>
        /// <strong>Hint To developers:</strong> Drew Varner &lt;Drew.Varner
        /// -at- sc.edu&gt; said that this is an image in WMF or Clipboard (BMP?)
        /// format. However, we won't do any conversion into any image type but
        /// instead just return a byte array.
        /// </summary>
        /// <value>The thumbnail.</value>
        public byte[] Thumbnail
        {
            get{return (byte[])GetProperty(PropertyIDMap.PID_THUMBNAIL);}
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_THUMBNAIL, /* FIXME: */
                        Variant.VT_LPSTR, value);
            }
        }

        /// <summary>
        /// Removes the thumbnail.
        /// </summary>
        public void RemoveThumbnail()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_THUMBNAIL);
        }


        /// <summary>
        /// Gets or sets the name of the application.
        /// </summary>
        /// <value>The name of the application.</value>
        public String ApplicationName
        {
            get { return GetPropertyStringValue(PropertyIDMap.PID_APPNAME); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_APPNAME, value);
            }
        }


        /// <summary>
        /// Removes the name of the application.
        /// </summary>
        public void RemoveApplicationName()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_APPNAME);
        }



        /// <summary>
        /// Gets or sets a security code which is one of the following values:
        /// <ul>
        /// 	<li>0 if the {@link SummaryInformation} does not contain a
        /// security field or if there is no security on the document. Use
        /// {@link PropertySet#wasNull()} To distinguish between the two
        /// cases!</li>
        /// 	<li>1 if the document is password protected</li>
        /// 	<li>2 if the document is Read-only recommended</li>
        /// 	<li>4 if the document is Read-only enforced</li>
        /// 	<li>8 if the document is locked for annotations</li>
        /// </ul>
        /// </summary>
        /// <value>The security code</value>
        public int Security
        {
            get { return GetPropertyIntValue(PropertyIDMap.PID_SECURITY); }
            set
            {
                MutableSection s = (MutableSection)FirstSection;
                s.SetProperty(PropertyIDMap.PID_SECURITY, value);
            }
        }


        /// <summary>
        /// Removes the security code.
        /// </summary>
        public void RemoveSecurity()
        {
            MutableSection s = (MutableSection)FirstSection;
            s.RemoveProperty(PropertyIDMap.PID_SECURITY);
        }

    }
}