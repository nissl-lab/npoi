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

namespace NPOI.XWPF.UserModel
{
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using NPOI.Util;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;

    public class XWPFComments : POIXMLDocumentPart
    {
        XWPFDocument document;
        private readonly List<XWPFComment> comments = new List<XWPFComment>();
        private readonly List<XWPFPictureData> pictures = new List<XWPFPictureData>();
        private CT_Comments ctComments;

        /**
         * Construct XWPFComments from a package part
         *
         * @param part the package part holding the data of the footnotes,
         */
        public XWPFComments(POIXMLDocumentPart parent, PackagePart part)
            : base(parent, part)
        {
            if (!(GetParent() is XWPFDocument))
            {
                throw new RuntimeException("Parent is not a XWPFDocuemnt: " + GetParent());
            }
            this.document = (XWPFDocument)GetParent();

            if (this.document == null)
            {
                throw new NullReferenceException();
            }
        }

        /**
         * Construct XWPFComments from scratch for a new document.
         */
        public XWPFComments()
        {
            ctComments = new CT_Comments(); //CT_Comments.Factory.newInstance();
        }

        /**
         * read comments form an existing package
         */
        internal override void OnDocumentRead()
        {
            try
            {
                using (var @is = GetPackagePart().GetInputStream())
                {
                    XmlDocument xmldoc = DocumentHelper.LoadDocument(@is);
                    CommentsDocument doc = CommentsDocument.Parse(xmldoc, NamespaceManager);
                    ctComments = doc.Comments;
                    foreach (CT_Comment ctComment in ctComments.comment)
                    {
                        comments.Add(new XWPFComment(ctComment, this));
                    }
                }
            }
            catch (XmlException e)
            {
                throw new POIXMLException("Unable to read comments", e);
            }

            foreach (POIXMLDocumentPart poixmlDocumentPart in GetRelations())
            {
                if (poixmlDocumentPart is XWPFPictureData)
                {
                    XWPFPictureData xwpfPicData = (XWPFPictureData)poixmlDocumentPart;
                    pictures.Add(xwpfPicData);
                    document.RegisterPackagePictureData(xwpfPicData);
                }
            }
        }

        /**
         * Adds a picture to the comments.
         *
         * @param is     The stream to read image from
         * @param format The format of the picture, see {@link Document}
         * @return the index to this picture (0 based), the added picture can be
         * obtained from {@link #getAllPictures()} .
         * @throws InvalidFormatException If the format of the picture is not known.
         * @throws IOException            If reading the picture-data from the stream fails.
         * @see #addPictureData(InputStream, PictureType)
         */
        public String AddPictureData(Stream @is, int format)
        {
            byte[] data = IOUtils.ToByteArray(@is);
            return AddPictureData(data, format);
        }

        ///**
        // * Adds a picture to the comments.
        // *
        // * @param is     The stream to read image from
        // * @param pictureType The {@link PictureType} of the picture
        // * @return the index to this picture (0 based), the added picture can be
        // * obtained from {@link #getAllPictures()} .
        // * @throws InvalidFormatException If the pictureType of the picture is not known.
        // * @throws IOException            If reading the picture-data from the stream fails.
        // * @since POI 5.2.3
        // */
        //public String AddPictureData(InputStream @is, PictureType pictureType)
        //{
        //    byte[] data = IOUtils.toByteArrayWithMaxLength(@is, XWPFPictureData.getMaxImageSize());
        //    return AddPictureData(data, pictureType);
        //}

        /**
         * Adds a picture to the comments.
         *
         * @param pictureData The picture data
         * @param format      The format of the picture, see {@link Document}
         * @return the index to this picture (0 based), the added picture can be
         * obtained from {@link #getAllPictures()} .
         * @throws InvalidFormatException If the format of the picture is not known.
         */
        public String AddPictureData(byte[] pictureData, int format)
        {
            XWPFPictureData xwpfPicData = document.FindPackagePictureData(pictureData, format);
            POIXMLRelation relDesc = XWPFPictureData.RELATIONS[format];

            if (xwpfPicData == null)
            {
                /* Part doesn't exist, create a new one */
                int idx = GetXWPFDocument().GetNextPicNameNumber(format);
                xwpfPicData = (XWPFPictureData)CreateRelationship(relDesc, XWPFFactory.GetInstance(), idx);
                /* write bytes to new part */
                PackagePart picDataPart = xwpfPicData.GetPackagePart();
                try
                {
                    using (Stream @out = picDataPart.GetOutputStream())
                    {
                        @out.Write(pictureData, 0, pictureData.Length);
                    }
                }
                catch (IOException e)
                {
                    throw new POIXMLException(e);
                }

                document.RegisterPackagePictureData(xwpfPicData);
                pictures.Add(xwpfPicData);
                return GetRelationId(xwpfPicData);
            }
            else if (!GetRelations().Contains(xwpfPicData))
            {
                /*
                 * Part already existed, but was not related so far. Create
                 * relationship to the already existing part and update
                 * POIXMLDocumentPart data.
                 */
                // TODO add support for TargetMode.EXTERNAL relations.
                RelationPart rp = AddRelation(null, XWPFRelation.IMAGES, xwpfPicData);
                pictures.Add(xwpfPicData);
                return rp.Relationship.Id;
            }
            else
            {
                /* Part already existed, get relation id and return it */
                return GetRelationId(xwpfPicData);
            }
        }

        /**
         * save and commit comments
         */
        protected internal override void Commit()
        {
            //XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            //xmlOptions.setSaveSyntheticDocumentElement(new QName(CTComments.type.getName().getNamespaceURI(), "comments"));

            PackagePart part = GetPackagePart();
            using (var @out = part.GetOutputStream())
            {
                var doc = new CommentsDocument(ctComments);
                doc.Save(@out);
            }
        }

        public IList<XWPFPictureData> GetAllPictures()
        {
            return pictures.AsReadOnly();
        }

        /**
         * Gets the underlying CTComments object for the comments.
         *
         * @return CTComments object
         */
        public CT_Comments GetCtComments()
        {
            return ctComments;
        }

        /**
         * set a new comments
         */
        //@Internal
        internal void SetCtComments(CT_Comments ctComments)
        {
            this.ctComments = ctComments;
        }

        /**
         * Get the list of {@link XWPFComment} in the Comments part.
         */
        public List<XWPFComment> GetComments()
        {
            return comments;
        }

        /**
         * Get the specified comment by position
         *
         * @param pos Array position of the comment
         */
        public XWPFComment GetComment(int pos)
        {
            if (pos >= 0 && pos < ctComments.comment.Count)
            {
                return GetComments()[pos];
            }
            return null;
        }

        /**
         * Get the specified comment by comment id
         *
         * @param id comment id
         * @return the specified comment
         */
        public XWPFComment GetCommentByID(String id)
        {
            foreach (XWPFComment comment in comments)
            {
                if (comment.Id.Equals(id))
                {
                    return comment;
                }
            }
            return null;
        }

        /**
         * Get the specified comment by ctComment
         */
        public XWPFComment GetComment(CT_Comment ctComment)
        {
            foreach (XWPFComment comment in comments)
            {
                if (comment.GetCtComment() == ctComment)
                {
                    return comment;
                }
            }
            return null;
        }

        /**
         * Create a new comment and add it to the document.
         *
         * @param cid comment Id
         */
        public XWPFComment CreateComment(string cid)
        {
            CT_Comment ctComment = ctComments.AddNewComment();
            ctComment.id = cid;
            XWPFComment comment = new XWPFComment(ctComment, this);
            comments.Add(comment);
            return comment;
        }

        /**
         * Remove the specified comment if present.
         *
         * @param pos Array position of the comment to be removed
         * @return True if the comment was removed.
         */
        public bool RemoveComment(int pos)
        {
            if (pos >= 0 && pos < ctComments.comment.Count)
            {
                comments.RemoveAt(pos);
                ctComments.RemoveComment(pos);
                return true;
            }
            return false;
        }

        public XWPFDocument GetXWPFDocument()
        {
            if (null != document)
            {
                return document;
            }
            return (XWPFDocument)GetParent();
        }

        public void SetXWPFDocument(XWPFDocument document)
        {
            this.document = document;
        }
    }
}
