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

namespace NPOI.SS.UserModel
{
    using System;

    public interface IPictureData
    {

        /// <summary>
        /// Gets the picture data.
        /// </summary>
        /// <returns>the picture data.</returns>
        byte[] Data { get; }

        /// <summary>
        /// Suggests a file extension for this image.
        /// </summary>
        /// <returns>the file extension.</returns>
        String SuggestFileExtension();
        /// <summary>
        /// Returns the mime type for the image
        /// </summary>
        String MimeType { get; }

        /// <summary>
        /// </summary>
        /// <returns>the POI internal image type, 0 if unknown image type</returns>
        /// 
        /// <see cref="Workbook.PICTURE_TYPE_DIB" />
        /// <see cref="Workbook.PICTURE_TYPE_EMF" />
        /// <see cref="Workbook.PICTURE_TYPE_JPEG" />
        /// <see cref="Workbook.PICTURE_TYPE_PICT" />
        /// <see cref="Workbook.PICTURE_TYPE_PNG" />
        /// <see cref="Workbook.PICTURE_TYPE_WMF" />
        PictureType PictureType { get; }
    }
}