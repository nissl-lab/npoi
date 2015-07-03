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

namespace NPOI.XWPF.UserModel
{
    using System;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using System.IO;

    /// <summary>
    /// Raw picture data, normally attached to a WordProcessingML Drawing. As a rule, pictures are stored in the /word/media/ part of a WordProcessingML package.
    /// </summary>
    /// <remarks>
    /// @author Philipp Epp
    /// </remarks>
    public class XWPFPictureData : POIXMLDocumentPart
    {

        /**
         * Relationships for each known picture type
         */
        internal static POIXMLRelation[] RELATIONS;
        static XWPFPictureData()
        {
            RELATIONS = new POIXMLRelation[13];
            RELATIONS[(int)PictureType.EMF] = XWPFRelation.IMAGE_EMF;
            RELATIONS[(int)PictureType.WMF] = XWPFRelation.IMAGE_WMF;
            RELATIONS[(int)PictureType.PICT] = XWPFRelation.IMAGE_PICT;
            RELATIONS[(int)PictureType.JPEG] = XWPFRelation.IMAGE_JPEG;
            RELATIONS[(int)PictureType.PNG] = XWPFRelation.IMAGE_PNG;
            RELATIONS[(int)PictureType.DIB] = XWPFRelation.IMAGE_DIB;
            RELATIONS[(int)PictureType.GIF] = XWPFRelation.IMAGE_GIF;
            RELATIONS[(int)PictureType.TIFF] = XWPFRelation.IMAGE_TIFF;
            RELATIONS[(int)PictureType.EPS] = XWPFRelation.IMAGE_EPS;
            RELATIONS[(int)PictureType.BMP] = XWPFRelation.IMAGE_BMP;
            RELATIONS[(int)PictureType.WPG] = XWPFRelation.IMAGE_WPG;
        }

        private long? checksum = null;

        /**
         * Create a new XWPFGraphicData node
         *
         */
        protected XWPFPictureData()
            : base()
        {
        }

        /**
         * Construct XWPFPictureData from a package part
         *
         * @param part the package part holding the Drawing data,
         * @param rel  the package relationship holding this Drawing,
         * the relationship type must be http://schemas.Openxmlformats.org/officeDocument/2006/relationships/image
         */
        public XWPFPictureData(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
        }


        internal override void OnDocumentRead()
        {
            base.OnDocumentRead();
        }

        /**
         * Gets the picture data as a byte array.
         * <p>
         * Note, that this call might be expensive since all the picture data is copied into a temporary byte array.
         * You can grab the picture data directly from the underlying package part as follows:
         * <br/>
         * <code>
         * InputStream is1 = GetPackagePart().InputStream;
         * </code>
         * </p>
         * @return the Picture data.
         */
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

        /**
         * Returns the file name of the image, eg image7.jpg . The original filename
         * isn't always available, but if it can be found it's likely to be in the
         * CTDrawing
         */
        public String FileName
        {
            get
            {
                String name = GetPackagePart().PartName.Name;
                if (name == null)
                    return null;
                return name.Substring(name.LastIndexOf('/') + 1);
            }
        }

        /**
         * Suggests a file extension for this image.
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
         * @see NPOI.XWPF.UserModel.PictureTypeEMF
         * @see NPOI.XWPF.UserModel.PictureTypeWMF
         * @see NPOI.XWPF.UserModel.PictureTypePICT
         * @see NPOI.XWPF.UserModel.PictureTypeJPEG
         * @see NPOI.XWPF.UserModel.PictureTypePNG
         * @see NPOI.XWPF.UserModel.PictureTypeDIB
         */
        public int GetPictureType()
        {
            String contentType = GetPackagePart().ContentType;
            for (int i = 0; i < RELATIONS.Length; i++)
            {
                if (RELATIONS[i] == null)
                {
                    continue;
                }

                if (RELATIONS[i].ContentType.Equals(contentType))
                {
                    return i;
                }
            }
            return 0;
        }

        public long Checksum
        {
            get
            {
                if (this.checksum == null)
                {
                    Stream is1 = null;
                    byte[] data;
                    try
                    {
                        is1 = GetPackagePart().GetInputStream();
                        data = IOUtils.ToByteArray(is1);
                    }
                    catch (IOException e)
                    {
                        throw new POIXMLException(e);
                    }
                    finally
                    {
                        try
                        {
                            if (is1 != null)
                                is1.Close();
                        }
                        catch (IOException e)
                        {
                            throw new POIXMLException(e);
                        }
                    }
                    this.checksum = IOUtils.CalculateChecksum(data);
                }
                return this.checksum.Value;
            }
        }


        public override bool Equals(Object obj)
        {
            /**
             * In case two objects ARE Equal, but its not the same instance, this
             * implementation will always run through the whole
             * byte-array-comparison before returning true. If this will turn into a
             * performance issue, two possible approaches are available:<br>
             * a) Use the Checksum only and take the risk that two images might have
             * the same CRC32 sum, although they are not the same.<br>
             * b) Use a second (or third) Checksum algorithm to minimise the chance
             * that two images have the same Checksums but are not equal (e.g.
             * CRC32, MD5 and SHA-1 Checksums, Additionally compare the
             * data-byte-array lengths).
             */
            if (obj == this)
            {
                return true;
            }

            if (obj == null)
            {
                return false;
            }

            if (!(obj is XWPFPictureData))
            {
                return false;
            }

            XWPFPictureData picData = (XWPFPictureData)obj;
            PackagePart foreignPackagePart = picData.GetPackagePart();
            PackagePart ownPackagePart = this.GetPackagePart();

            if ((foreignPackagePart != null && ownPackagePart == null)
                    || (foreignPackagePart == null && ownPackagePart != null))
            {
                return false;
            }

            if (ownPackagePart != null)
            {
                OPCPackage foreignPackage = foreignPackagePart.Package;
                OPCPackage ownPackage = ownPackagePart.Package;

                if ((foreignPackage != null && ownPackage == null)
                        || (foreignPackage == null && ownPackage != null))
                {
                    return false;
                }
                if (ownPackage != null)
                {

                    if (!ownPackage.Equals(foreignPackage))
                    {
                        return false;
                    }
                }
            }

            long foreignChecksum = picData.Checksum;
            long localChecksum = Checksum;

            if (!(localChecksum.Equals(foreignChecksum)))
            {
                return false;
            }
            return Arrays.Equals(this.Data, picData.Data);
        }


        public override int GetHashCode()
        {
            return Checksum.GetHashCode();
        }

        /**
         * *PictureData objects store the actual content in the part directly without keeping a 
         * copy like all others therefore we need to handle them differently.
         */
        protected internal override void PrepareForCommit()
        {
            // do not clear the part here
        }
    }

}