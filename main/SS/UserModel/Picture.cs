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

using SixLabors.ImageSharp;
namespace NPOI.SS.UserModel
{
    public enum PictureType : int
    {
        /// <summary>
        /// Allow accessing the Initial value.
        /// </summary>
        /// 
        Unknown = -1,

        None = 0,

        /// <summary>
        /// Extended windows meta file */
        /// </summary>
        EMF = 2,

        /// <summary>
        /// Windows Meta File */
        /// </summary>
        WMF = 3,

        /// <summary>
        /// Mac PICT format */
        /// </summary>
        PICT = 4,

        /// <summary>
        /// JPEG format */
        /// </summary>
        JPEG = 5,

        /// <summary>
        /// PNG format */
        /// </summary>
        PNG = 6,

        /// <summary>
        /// Device independent bitmap */
        /// </summary>
        DIB = 7,

        /// <summary>
        /// GIF image format */
        /// </summary>
        GIF = 8,
        /// <summary>
        /// Tag Image File (.tiff)
        /// </summary>
        TIFF = 9,

        /// <summary>
        /// Encapsulated Postscript (.eps)
        /// </summary>
        EPS = 10,


        /// <summary>
        /// Windows Bitmap (.bmp)
        /// </summary>
        BMP = 11,

        /// <summary>
        /// WordPerfect graphics (.wpg)
        /// </summary>
        WPG = 12
    }
    /// <summary>
    /// Repersents a picture in a SpreadsheetML document
    /// </summary>
    /// @author Yegor Kozlov
    public interface IPicture : IShape
    {

        /// <summary>
        /// Reset the image to the dimension of the embedded image
        /// </summary>
        /// <see cref="resize(double, double)" />
        void Resize();

        /// <summary>
        /// Resize the image proportionally.
        /// </summary>
        void Resize(double scale);

        /// <summary>
        /// <para>
        /// Resize the image.
        /// </para>
        /// <para>
        /// Please note, that this method works correctly only for workbooks
        /// with the default font size (Arial 10pt for .xls and Calibri 11pt for .xlsx).
        /// If the default font is changed the resized image can be streched vertically or horizontally.
        /// </para>
        /// <para>
        /// 
        /// </para>
        /// <para>
        /// <c>resize(1.0,1.0)</c> keeps the original size,<br/>
        /// <c>resize(0.5,0.5)</c> resize to 50% of the original,<br/>
        /// <c>resize(2.0,2.0)</c> resizes to 200% of the original.<br/>
        /// <c>resize(<see cref="Double.MAX_VALUE" />,<see cref="Double.MAX_VALUE" />)</c> resizes to the dimension of the embedded image.
        /// </para>
        /// </summary>
        /// <param name="scaleX">the amount by which the image width is multiplied relative to the original width.</param>
        /// <param name="scaleY">the amount by which the image height is multiplied relative to the original height.</param>
        void Resize(double scaleX, double scaleY);

        /// <summary>
        /// Calculate the preferred size for this picture.
        /// </summary>
        /// <returns>XSSFClientAnchor with the preferred size for this image</returns>
        IClientAnchor GetPreferredSize();

        /// <summary>
        /// Calculate the preferred size for this picture.
        /// </summary>
        /// <param name="scaleX">the amount by which image width is multiplied relative to the original width.</param>
        /// <param name="scaleY">the amount by which image height is multiplied relative to the original height.</param>
        /// <returns>ClientAnchor with the preferred size for this image</returns>
        IClientAnchor GetPreferredSize(double scaleX, double scaleY);

        /// <summary>
        /// Return the dimension of the embedded image in pixel
        /// </summary>
        /// <returns>image dimension in pixels</returns>
        Size GetImageDimension();



        /// <summary>
        /// Return picture data for this picture
        /// </summary>
        /// <returns>picture data for this picture</returns>
        IPictureData PictureData { get; }
        /// <summary>
        /// </summary>
        /// <returns>the anchor that is used by this picture</returns>
        IClientAnchor ClientAnchor { get; }
        /*
         * @return the sheet which contains the picture
         */
        ISheet Sheet { get; }
    }
}
