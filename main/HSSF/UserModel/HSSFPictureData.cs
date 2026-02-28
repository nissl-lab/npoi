/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed Under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.DDF;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    /// <summary>
    /// Represents binary data stored in the file.  Eg. A GIF, JPEG etc...
    /// @author Daniel Noll
    /// </summary>
    public class HSSFPictureData : IPictureData
    {
        // MSOBI constants for various formats.
        public const short MSOBI_WMF = 0x2160;
        public const short MSOBI_EMF = 0x3D40;
        public const short MSOBI_PICT = 0x5420;
        public const short MSOBI_PNG = 0x6E00;
        public const short MSOBI_JPEG = 0x46A0;
        public const short MSOBI_DIB = 0x7A80;
        // Mask of the bits in the options used to store the image format.
        public const short FORMAT_MASK = unchecked((short)0xFFF0);

        /**
         * Underlying escher blip record containing the bitmap data.
         */
        private EscherBlipRecord blip;

        /// <summary>
        /// Constructs a picture object.
        /// </summary>
        /// <param name="blip">the underlying blip record containing the bitmap data.</param>
        public HSSFPictureData(EscherBlipRecord blip)
        {
            this.blip = blip;
        }

        /// <summary>
        /// Gets the picture data.
        /// </summary>
        /// <value>the picture data.</value>
        public byte[] Data
        {
            get
            {
                byte[] pictureData = blip.PictureData;

                //PNG created on MAC may have a 16-byte prefix which prevents successful reading.
                //Just cut it off!.
                if (PngUtils.MatchesPngHeader(pictureData, 16))
                {
                    byte[] png = new byte[pictureData.Length - 16];
                    System.Array.Copy(pictureData, 16, png, 0, png.Length);
                    pictureData = png;
                }

                return pictureData;
            }
        }
        /// <summary>
        /// gets format of the picture.
        /// </summary>
        /// <value>The format.</value>
        public int Format
        {
            get
            {
                return blip.RecordId - unchecked((short)0xF018);
            }
        }
        /// <summary>
        /// Suggests a file extension for this image.
        /// </summary>
        /// <returns>the file extension.</returns>
        public String SuggestFileExtension()
        {
            switch (blip.RecordId)
            {
                case EscherMetafileBlip.RECORD_ID_WMF:
                    return "wmf";
                case EscherMetafileBlip.RECORD_ID_EMF:
                    return "emf";
                case EscherMetafileBlip.RECORD_ID_PICT:
                    return "pict";
                case EscherBitmapBlip.RECORD_ID_PNG:
                    return "png";
                case EscherBitmapBlip.RECORD_ID_JPEG:
                    return "jpeg";
                case EscherBitmapBlip.RECORD_ID_DIB:
                    return "dib";
                default:
                    return "";
            }
        }
        /**
     * Returns the mime type for the image
     */
        public String MimeType
        {
            get
            {
                switch (blip.RecordId)
                {
                    case EscherMetafileBlip.RECORD_ID_WMF:
                        return "image/x-wmf";
                    case EscherMetafileBlip.RECORD_ID_EMF:
                        return "image/x-emf";
                    case EscherMetafileBlip.RECORD_ID_PICT:
                        return "image/x-pict";
                    case EscherBitmapBlip.RECORD_ID_PNG:
                        return "image/png";
                    case EscherBitmapBlip.RECORD_ID_JPEG:
                        return "image/jpeg";
                    case EscherBitmapBlip.RECORD_ID_DIB:
                        return "image/bmp";
                    default:
                        return "image/unknown";
                }
            }
        }

        /**
     * @return the POI internal image type, -1 if not unknown image type
     *
     * @see Workbook#PICTURE_TYPE_DIB
     * @see Workbook#PICTURE_TYPE_EMF
     * @see Workbook#PICTURE_TYPE_JPEG
     * @see Workbook#PICTURE_TYPE_PICT
     * @see Workbook#PICTURE_TYPE_PNG
     * @see Workbook#PICTURE_TYPE_WMF
     */
        public PictureType PictureType
        {
            get
            {
                switch (blip.RecordId)
                {
                    case EscherMetafileBlip.RECORD_ID_WMF:
                        return PictureType.WMF;
                    case EscherMetafileBlip.RECORD_ID_EMF:
                        return PictureType.EMF;
                    case EscherMetafileBlip.RECORD_ID_PICT:
                        return PictureType.PICT;
                    case EscherBitmapBlip.RECORD_ID_PNG:
                        return PictureType.PNG;
                    case EscherBitmapBlip.RECORD_ID_JPEG:
                        return PictureType.JPEG;
                    case EscherBitmapBlip.RECORD_ID_DIB:
                        return PictureType.DIB;
                    default:
                        return PictureType.Unknown;
                }
            }
        }
    }
}
