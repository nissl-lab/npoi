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
namespace NPOI.XSSF.Model
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.XSSF.UserModel;
    using OpenXmlFormats.Spreadsheet;

    public class CommentsTable : POIXMLDocumentPart
    {
        private CT_Comments comments;
        /**
         * XML Beans uses a list, which is very slow
         *  to search, so we wrap things with our own
         *  map for fast Lookup.
         */
        private Dictionary<String, CT_Comment> commentRefs;

        public CommentsTable()
            : base()
        {

            comments = new CT_Comments();
            comments.AddNewCommentList();
            comments.AddNewAuthors().AddAuthor("");
        }

        internal CommentsTable(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {

            XmlDocument xml = ConvertStreamToXml(part.GetInputStream());
            ReadFrom(xml);
        }

        public void ReadFrom(XmlDocument xmlDoc)
        {
            try
            {
                CommentsDocument doc = CommentsDocument.Parse(xmlDoc, NamespaceManager);
                comments = doc.GetComments();

            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }
        public void WriteTo(Stream out1)
        {
            CommentsDocument doc = new CommentsDocument();
            doc.SetComments(comments);
            doc.Save(out1);
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }

        /**
         * Called after the reference is updated, so that
         *  we can reflect that in our cache
         */
        public void ReferenceUpdated(String oldReference, CT_Comment comment)
        {
            if (commentRefs != null)
            {
                commentRefs.Remove(oldReference);
                commentRefs[comment.@ref] = comment;
            }
        }

        public void RecreateReference()
        {
            commentRefs.Clear();
            foreach (CT_Comment comment in comments.commentList.comment)
            {
                commentRefs.Add(comment.@ref, comment);
            }
        }

        public int GetNumberOfComments()
        {
            return comments.commentList.SizeOfCommentArray();
        }

        public int GetNumberOfAuthors()
        {
            return comments.authors.SizeOfAuthorArray();
        }

        public String GetAuthor(long authorId)
        {
            return comments.authors.GetAuthorArray((int)authorId);
        }

        /// <summary>
        /// Searches the author. If not found he is added to the list of authors.
        /// </summary>
        /// <param name="author">author to search</param>
        /// <returns>index of the author</returns>
        public int FindAuthor(String author)
        {
            for (int i = 0; i < comments.authors.SizeOfAuthorArray(); i++)
            {
                if (comments.authors.GetAuthorArray(i).Equals(author))
                {
                    return i;
                }
            }
            return AddNewAuthor(author);
        }

        public XSSFComment FindCellComment(String cellRef)
        {
            CT_Comment ct = GetCTComment(cellRef);
            return ct == null ? null : new XSSFComment(this, ct, null);
        }

        public CT_Comment GetCTComment(String cellRef)
        {
            // Create the cache if needed
            if (commentRefs == null)
            {
                commentRefs = new Dictionary<String, CT_Comment>();
                if (comments.commentList.comment != null)
                {
                    foreach (CT_Comment comment in comments.commentList.comment)
                    {
                        commentRefs.Add(comment.@ref, comment);
                    }
                }
            }

            // Return the comment, or null if not known
            if (!commentRefs.ContainsKey(cellRef))
                return null;
            return commentRefs[cellRef];
        }

        /**
         * This method is deprecated and should not be used any more as
         * it overwrites the comment in Cell A1.
         *
         * @return
         */
        [Obsolete]
        public CT_Comment NewComment() {
            return NewComment("A1");
        }

        public CT_Comment NewComment(String ref1)
        {
            CT_Comment ct = comments.commentList.AddNewComment();
            ct.@ref = (ref1);
            ct.authorId = (0);

            if (commentRefs != null)
            {
                commentRefs[ct.@ref] = ct;
            }
            return ct;
        }

        public bool RemoveComment(String cellRef)
        {
            CT_CommentList lst = comments.commentList;
            if (lst != null) for (int i = 0; i < lst.SizeOfCommentArray(); i++)
                {
                    CT_Comment comment = lst.GetCommentArray(i);
                    if (cellRef.Equals(comment.@ref))
                    {
                        lst.RemoveComment(i);

                        if (commentRefs != null)
                        {
                            commentRefs.Remove(cellRef);
                        }
                        return true;
                    }
                }
            return false;
        }

        private int AddNewAuthor(String author)
        {
            int index = comments.authors.SizeOfAuthorArray();
            comments.authors.Insert(index, author);
            return index;
        }

        public CT_Comments GetCTComments()
        {
            return comments;
        }
    }

}



