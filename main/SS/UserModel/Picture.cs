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
    public enum PictureType : int
    {
        /// <summary>
        /// Allow accessing the Initial value.
        /// </summary>
        None = 0,
        /** Extended windows meta file */
        EMF = 2,
        /** Windows Meta File */
        WMF = 3,
        /** Mac PICT format */
        PICT = 4,
        /** JPEG format */
        JPEG = 5,
        /** PNG format */
        PNG = 6,
        /** Device independant bitmap */
        DIB = 7,
    }
    /**
     * Repersents a picture in a SpreadsheetML document
     *
     * @author Yegor Kozlov
     */
    public interface IPicture
    {

        /**
         * Reset the image to the original size.
         */
        void Resize();

        /**
         * Reset the image to the original size.
         *
         * @param scale the amount by which image dimensions are multiplied relative to the original size.
         * <code>resize(1.0)</code> Sets the original size, <code>resize(0.5)</code> resize to 50% of the original,
         * <code>resize(2.0)</code> resizes to 200% of the original.
         */
        void Resize(double scale);

        IClientAnchor GetPreferredSize();
        /**
         * Return picture data for this picture
         *
         * @return picture data for this picture
         */
        IPictureData PictureData { get; }
    }
}
