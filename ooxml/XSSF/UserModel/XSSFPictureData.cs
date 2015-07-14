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

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.Util;
using System.IO;
using System;
using System.Collections.Generic;
namespace NPOI.XSSF.UserModel
{

    /**
     * Raw picture data, normally attached to a SpreadsheetML Drawing.
     * As a rule, pictures are stored in the /xl/media/ part of a SpreadsheetML package.
     */
    public class XSSFPictureData : POIXMLDocumentPart, IPictureData
    {

        /**
         * Relationships for each known picture type
         */
        internal static Dictionary<int, POIXMLRelation> RELATIONS;
        static XSSFPictureData()
        {
            RELATIONS = new Dictionary<int,POIXMLRelation>(8);
            RELATIONS[(int)PictureType.EMF] = XSSFRelation.IMAGE_EMF;
            RELATIONS[(int)PictureType.WMF] = XSSFRelation.IMAGE_WMF;
            RELATIONS[(int)PictureType.PICT] = XSSFRelation.IMAGE_PICT;
            RELATIONS[(int)PictureType.JPEG] = XSSFRelation.IMAGE_JPEG;
            RELATIONS[(int)PictureType.PNG] = XSSFRelation.IMAGE_PNG;
            RELATIONS[(int)PictureType.DIB] = XSSFRelation.IMAGE_DIB;
            RELATIONS[XSSFWorkbook.PICTURE_TYPE_GIF] = XSSFRelation.IMAGE_GIF;
            RELATIONS[XSSFWorkbook.PICTURE_TYPE_TIFF] = XSSFRelation.IMAGE_TIFF;
            RELATIONS[XSSFWorkbook.PICTURE_TYPE_EPS] = XSSFRelation.IMAGE_EPS;
            RELATIONS[XSSFWorkbook.PICTURE_TYPE_BMP] = XSSFRelation.IMAGE_BMP;
            RELATIONS[XSSFWorkbook.PICTURE_TYPE_WPG] = XSSFRelation.IMAGE_WPG;
        }

        /**
         * Create a new XSSFPictureData node
         *
         * @see NPOI.xssf.usermodel.XSSFWorkbook#AddPicture(byte[], int)
         */
        public XSSFPictureData()
            : base()
        {

        }

        /**
         * Construct XSSFPictureData from a namespace part
         *
         * @param part the namespace part holding the Drawing data,
         * @param rel  the namespace relationship holding this Drawing,
         * the relationship type must be http://schemas.Openxmlformats.org/officeDocument/2006/relationships/image
         */
        internal XSSFPictureData(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

        }


        /**
         * Suggests a file extension for this image.
         *
         * @return the file extension.
         */
        public String SuggestFileExtension()
        {
            return GetPackagePart().PartName.Extension;
        }

        /**
         * Return an integer constant that specifies type of this picture
         *
         * @return an integer constant that specifies type of this picture 
         * @see NPOI.ss.usermodel.Workbook#PICTURE_TYPE_EMF
         * @see NPOI.ss.usermodel.Workbook#PICTURE_TYPE_WMF
         * @see NPOI.ss.usermodel.Workbook#PICTURE_TYPE_PICT
         * @see NPOI.ss.usermodel.Workbook#PICTURE_TYPE_JPEG
         * @see NPOI.ss.usermodel.Workbook#PICTURE_TYPE_PNG
         * @see NPOI.ss.usermodel.Workbook#PICTURE_TYPE_DIB
         */
        public PictureType PictureType
        {
            get
            {
                String contentType = GetPackagePart().ContentType;
                foreach (PictureType relation in RELATIONS.Keys)
                {
                    if (RELATIONS[(int)relation].ContentType.Equals(contentType))
                    {
                        return relation;
                    }
                }
                return PictureType.None;
            }
        }


        /// <summary>
        /// Gets the picture data as a byte array.
        /// </summary>
        public byte[] Data
        {
            get
            {
                try
                {
                    return IOUtils.ToByteArray(GetPackagePart().GetInputStream());
                }
                catch (IOException e)
                {
                    throw new POIXMLException(e);
                }
            }
        }

        public string MimeType
        {
            get { return GetPackagePart().ContentType; }
        }

        /**
         * *PictureData objects store the actual content in the part directly without keeping a 
         * copy like all others therefore we need to handle them differently.
         */
        protected internal override void PrepareForCommit() {
            // do not clear the part here
        }
    }
}


