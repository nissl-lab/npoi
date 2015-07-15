/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
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
    using NPOI.Util;

    /// <summary>
    /// Class To manipulate data in the Clipboard Variant (Variant#VT_CF VT_CF) format.
    /// @author Drew Varner (Drew.Varner inOrAround sc.edu)
    /// @since 2002-04-29
    /// </summary>
    public class Thumbnail
    {
        /// <summary>
        /// OffSet in bytes where the Clipboard Format Tag starts in the <c>byte[]</c> returned by SummaryInformation#GetThumbnail()
        /// </summary>
        public const int OFFSet_CFTAG = 4;

        /// <summary>
        /// OffSet in bytes where the Clipboard Format starts in the <c>byte[]</c> returned by SummaryInformation#GetThumbnail()
        /// </summary>
        /// <remarks>This is only valid if the Clipboard Format Tag is CFTAG_WINDOWS</remarks>
        public const int OFFSet_CF = 8;

        /// <summary>
        /// OffSet in bytes where the Windows Metafile (WMF) image data starts in the <c>byte[]</c> returned by SummaryInformation#GetThumbnail()
        /// There is only WMF data at this point in the
        /// <c>byte[]</c> if the Clipboard Format Tag is
        /// CFTAG_WINDOWS and the Clipboard Format is 
        /// CF_METAFILEPICT.
        /// </summary>
        /// <remarks>Note: The <c>byte[]</c> that starts at
        /// <c>OFFSet_WMFDATA</c> and ends at
        /// <c>GetThumbnail().Length - 1</c> forms a complete WMF
        /// image. It can be saved To disk with a <c>.wmf</c> file
        /// type and Read using a WMF-capable image viewer.</remarks>
        public const int OFFSet_WMFDATA = 20;

        /// <summary>
        /// Clipboard Format Tag - Windows clipboard format
        /// </summary>
        /// <remarks>A <c>DWORD</c> indicating a built-in Windows clipboard format value</remarks>
        public const int CFTAG_WINDOWS = -1;

        /// <summary>
        /// Clipboard Format Tag - Macintosh clipboard format
        /// </summary>
        /// <remarks>A <c>DWORD</c> indicating a Macintosh clipboard format value</remarks>
        public const int CFTAG_MACINTOSH = -2;

        /// <summary>
        /// Clipboard Format Tag - Format ID
        /// </summary>
        /// <remarks>A GUID containing a format identifier (FMTID). This is rarely used.</remarks>
        public const int CFTAG_FMTID = -3;

        /// <summary>
        /// Clipboard Format Tag - No Data
        /// </summary>
        /// <remarks>A <c>DWORD</c> indicating No data. This is rarely used.</remarks>
        public const int CFTAG_NODATA = 0;

        /// <summary>
        /// Clipboard Format - Windows metafile format. This is the recommended way To store thumbnails in Property Streams.
        /// </summary>
        /// <remarks>Note:This is not the same format used in
        /// regular WMF images. The clipboard version of this format has an
        /// extra clipboard-specific header.</remarks>
        public const int CF_METAFILEPICT = 3;

        /// <summary>
        /// Clipboard Format - Device Independent Bitmap
        /// </summary>
        public const int CF_DIB = 8;
        /// <summary>
        /// Clipboard Format - Enhanced Windows metafile format
        /// </summary>
        public const int CF_ENHMETAFILE = 14;



        /// <summary>
        /// Clipboard Format - Bitmap
        /// </summary>
        /// <remarks>see msdn.microsoft.com/library/en-us/dnw98bk/html/clipboardoperations.asp</remarks>
        [Obsolete]
        public const int CF_BITMAP = 2;

        /**
         * A <c>byte[]</c> To hold a thumbnail image in (
         * Variant#VT_CF VT_CF) format.
         */
        private byte[] thumbnailData = null;



        /// <summary>
        /// Default Constructor. If you use it then one you'll have To Add
        /// the thumbnail <c>byte[]</c> from {@link
        /// SummaryInformation#GetThumbnail()} To do any useful
        /// manipulations, otherwise you'll Get a
        /// <c>NullPointerException</c>.
        /// </summary>
        public Thumbnail()
        {
            
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="Thumbnail"/> class.
        /// </summary>
        /// <param name="thumbnailData">The thumbnail data.</param>
        public Thumbnail(byte[] thumbnailData)
        {
            this.thumbnailData = thumbnailData;
        }



        /// <summary>
        /// Gets or sets the thumbnail as a <c>byte[]</c> in {@link
        /// Variant#VT_CF VT_CF} format.
        /// </summary>
        /// <value>The thumbnail value</value>
        public byte[] ThumbnailData
        {
            get
            {
                return thumbnailData;
            }
            set { this.thumbnailData = value; }
        }


        /// <summary>
        /// Returns an <c>int</c> representing the Clipboard
        /// Format Tag
        /// Possible return values are:
        /// <ul>
        /// 	<li>{@link #CFTAG_WINDOWS CFTAG_WINDOWS}</li>
        /// 	<li>{@link #CFTAG_MACINTOSH CFTAG_MACINTOSH}</li>
        /// 	<li>{@link #CFTAG_FMTID CFTAG_FMTID}</li>
        /// 	<li>{@link #CFTAG_NODATA CFTAG_NODATA}</li>
        /// </ul>
        /// </summary>
        /// <returns>A flag indicating the Clipboard Format Tag</returns>
        public long ClipboardFormatTag
        {
            get
            {
                long clipboardFormatTag = LittleEndian.GetInt(this.ThumbnailData,
                                                               OFFSet_CFTAG);
                return clipboardFormatTag;
            }
        }



        /// <summary>
        /// Returns an <c>int</c> representing the Clipboard
        /// Format
        /// Will throw an exception if the Thumbnail's Clipboard Format
        /// Tag is not {@link Thumbnail#CFTAG_WINDOWS CFTAG_WINDOWS}.
        /// Possible return values are:
        /// <ul>
        /// 	<li>{@link #CF_METAFILEPICT CF_METAFILEPICT}</li>
        /// 	<li>{@link #CF_DIB CF_DIB}</li>
        /// 	<li>{@link #CF_ENHMETAFILE CF_ENHMETAFILE}</li>
        /// 	<li>{@link #CF_BITMAP CF_BITMAP}</li>
        /// </ul>
        /// </summary>
        /// <returns>a flag indicating the Clipboard Format</returns>
        public long GetClipboardFormat()
        {
            if (!(ClipboardFormatTag == CFTAG_WINDOWS))
                throw new HPSFException("Clipboard Format Tag of Thumbnail must " +
                                        "be CFTAG_WINDOWS.");

            return LittleEndian.GetInt(this.ThumbnailData, OFFSet_CF);
        }



        /// <summary>
        /// Returns the Thumbnail as a <c>byte[]</c> of WMF data
        /// if the Thumbnail's Clipboard Format Tag is {@link
        /// #CFTAG_WINDOWS CFTAG_WINDOWS} and its Clipboard Format is
        /// {@link #CF_METAFILEPICT CF_METAFILEPICT}
        /// This
        /// <c>byte[]</c> is in the traditional WMF file, not the
        /// clipboard-specific version with special headers.
        /// See <a href="http://www.wvware.com/caolan/ora-wmf.html" tarGet="_blank">http://www.wvware.com/caolan/ora-wmf.html</a>
        /// for more information on the WMF image format.
        /// @return A WMF image of the Thumbnail
        /// @throws HPSFException if the Thumbnail isn't CFTAG_WINDOWS and
        /// CF_METAFILEPICT
        /// </summary>
        /// <returns></returns>
        public byte[] GetThumbnailAsWMF()
        {
            if (!(ClipboardFormatTag == CFTAG_WINDOWS))
                throw new HPSFException("Clipboard Format Tag of Thumbnail must " +
                                        "be CFTAG_WINDOWS.");
            if (!(GetClipboardFormat() == CF_METAFILEPICT))
                throw new HPSFException("Clipboard Format of Thumbnail must " +
                                        "be CF_METAFILEPICT.");
            else
            {
                byte[] thumbnail = this.ThumbnailData;
                int wmfImageLength = thumbnail.Length - OFFSet_WMFDATA;
                byte[] wmfImage = new byte[wmfImageLength];
                System.Array.Copy(thumbnail,
                                 OFFSet_WMFDATA,
                                 wmfImage,
                                 0,
                                 wmfImageLength);
                return wmfImage;
            }
        }

    }
}