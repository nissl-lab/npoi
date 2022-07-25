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
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;
    using NPOI.OpenXmlFormats.Spreadsheet;

    public class CommentsTable : POIXMLDocumentPart
    {
        public static String DEFAULT_AUTHOR = "";
        public static int DEFAULT_AUTHOR_ID = 0;
        private CT_Comments comments;
        /**
         * XML Beans uses a list, which is very slow
         *  to search, so we wrap things with our own
         *  map for fast Lookup.
         */
        private Dictionary<CellAddress, CT_Comment> commentRefs;

        public CommentsTable()
            : base()
        {

            comments = new CT_Comments();
            comments.AddNewCommentList();
            comments.AddNewAuthors().AddAuthor(DEFAULT_AUTHOR);
        }

        internal CommentsTable(PackagePart part)
            : base(part)
        {
            ReadFrom(part.GetInputStream());
        }

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public CommentsTable(PackagePart part, PackageRelationship rel)
             : this(part)
        {

        }

        public void ReadFrom(Stream is1)
        {
            try
            {
                XmlDocument xmlDoc = ConvertStreamToXml(is1);
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
        [Obsolete("2015-11-23 (circa POI 3.14beta1). Use {@link #referenceUpdated(CellAddress, CTComment)} instead")]
        public void ReferenceUpdated(String oldReference, CT_Comment comment)
        {
            ReferenceUpdated(new CellAddress(oldReference), comment);
        }

        /**
         * Called after the reference is updated, so that
         *  we can reflect that in our cache
         *  @param oldReference the comment to remove from the commentRefs map
         *  @param comment the comment to replace in the commentRefs map
         */
        public void ReferenceUpdated(CellAddress oldReference, CT_Comment comment)
        {
            if (commentRefs != null)
            {
                commentRefs.Remove(oldReference);
                commentRefs[new CellAddress(comment.@ref)] = comment;
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

        /**
         * Finds the cell comment at cellAddress, if one exists
         *
         * @param cellAddress the address of the cell to find a comment
         * @return cell comment if one exists, otherwise returns null
         * @
         */
        [Obsolete("deprecated 2015-11-23 (circa POI 3.14beta1). Use {@link #findCellComment(CellAddress)} instead")]
        public XSSFComment FindCellComment(String cellRef)
        {
            return FindCellComment(new CellAddress(cellRef));
        }

        /**
         * Finds the cell comment at cellAddress, if one exists
         *
         * @param cellAddress the address of the cell to find a comment
         * @return cell comment if one exists, otherwise returns null
         */
        public XSSFComment FindCellComment(CellAddress cellAddress)
        {
            CT_Comment ct = GetCTComment(cellAddress);
            return ct == null ? null : new XSSFComment(this, ct, null);
        }


        /**
         * Get the underlying CTComment xmlbean for a comment located at cellRef, if it exists
         *
         * @param cellRef the location of the cell comment
         * @return CTComment xmlbean if comment exists, otherwise return null.
         * @
         */
         [Obsolete("deprecated 2015-11-23 (circa POI 3.14beta1). Use {@link CommentsTable#getCTComment(CellAddress)} instead")]
        public CT_Comment GetCTComment(String ref1)
        {
            prepareCTCommentCache();
            return GetCTComment(new CellAddress(ref1));
        }

        public List<CellAddress> GetCellAddresses()
        {
            prepareCTCommentCache();
            return new List<CellAddress>(commentRefs.Keys);
        }

        /**
         * Get the underlying CTComment xmlbean for a comment located at cellRef, if it exists
         *
         * @param cellRef the location of the cell comment
         * @return CTComment xmlbean if comment exists, otherwise return null.
         */
        public CT_Comment GetCTComment(CellAddress cellRef)
        {
            // Create the cache if needed
            PrepareCTCommentCache();

            // Return the comment, or null if not known
            if (commentRefs.ContainsKey(cellRef))
                return commentRefs[cellRef];
            return null;
        }

        /**
         * Returns all cell comments on this sheet.
         * @return A map of each Comment in this sheet, keyed on the cell address where
         * the comment is located.
         */
        public Dictionary<CellAddress, IComment> GetCellComments()
        {
            PrepareCTCommentCache();
            Dictionary<CellAddress, IComment> map = new Dictionary<CellAddress, IComment>();

            foreach (KeyValuePair< CellAddress, CT_Comment> e in commentRefs)
            {
                map.Add(e.Key, new XSSFComment(this, e.Value, null));
            }

            return map;
        }

        /**
         * Refresh Map&lt;CellAddress, CTComment&gt; commentRefs cache,
         * Calls that use the commentRefs cache will perform in O(1)
         * time rather than O(n) lookup time for List<CTComment> comments.
         */
        private void PrepareCTCommentCache()
        {
            // Create the cache if needed
            if (commentRefs == null)
            {
                commentRefs = new Dictionary<CellAddress, CT_Comment>();
                foreach (CT_Comment comment in comments.commentList.GetCommentArray())
                {
                    commentRefs.Add(new CellAddress(comment.@ref), comment);
                }
            }
        }

        /**
         * Create a new comment located at cell address
         *
         * @param ref the location to add the comment
         * @return a new CTComment located at ref with default author
         */
        [Obsolete("deprecated 2015-11-23 (circa POI 3.14beta1). Use {@link #newComment(CellAddress)} instead")]
        public CT_Comment NewComment(String ref1)
        {
            return NewComment(new CellAddress(ref1));
        }

        /**
         * Create a new comment located` at cell address
         *
         * @param ref the location to add the comment
         * @return a new CTComment located at ref with default author
         */
        public CT_Comment NewComment(CellAddress ref1)
        {
            CT_Comment ct = comments.commentList.AddNewComment();
            ct.@ref = ref1.FormatAsString();
            ct.authorId = (uint)DEFAULT_AUTHOR_ID;

            if (commentRefs != null)
            {
                commentRefs.Add(ref1, ct);
            }
            return ct;
        }

        /**
         * Remove the comment at cellRef location, if one exists
         *
         * @param cellRef the location of the comment to remove
         * @return returns true if a comment was removed
         * @deprecated 2015-11-23 (circa POI 3.14beta1). Use {@link #removeComment(CellAddress)} instead
         */
        [Obsolete("deprecated 2015-11-23 (circa POI 3.14beta1). Use {@link #removeComment(CellAddress)} instead")]
        public bool RemoveComment(String cellRef)
        {
            return RemoveComment(new CellAddress(cellRef));
        }

        /**
         * Remove the comment at cellRef location, if one exists
         *
         * @param cellRef the location of the comment to remove
         * @return returns true if a comment was removed
         */
        public bool RemoveComment(CellAddress cellRef)
        {
            String stringRef = cellRef.FormatAsString();
            CT_CommentList lst = comments.commentList;
            if (lst != null)
            {
                CT_Comment[] commentArray = lst.GetCommentArray();
                for (int i = 0; i < commentArray.Length; i++)
                {
                    CT_Comment comment = commentArray[i];
                    if (stringRef.Equals(comment.@ref))
                    {
                        lst.RemoveComment(i);

                        if (commentRefs != null)
                        {
                            commentRefs.Remove(cellRef);
                        }
                        return true;
                    }
                }
            }
            return false;
        }

        /**
         * Add a new author to the CommentsTable.
         * This does not check if the author already exists.
         *
         * @param author the name of the comment author
         * @return the index of the new author
         */
        private int AddNewAuthor(String author)
        {
            int index = comments.authors.SizeOfAuthorArray();
            comments.authors.Insert(index, author);
            return index;
        }
        /**
         * Returns the underlying CTComments list xmlbean
         *
         * @return underlying comments list xmlbean
         */
        public CT_Comments GetCTComments()
        {
            return comments;
        }

        private void prepareCTCommentCache()
        {
            // Create the cache if needed
            if (commentRefs == null)
            {
                commentRefs = new Dictionary<CellAddress, CT_Comment>();
                foreach (CT_Comment comment in comments.commentList.GetCommentArray())
                {
                    commentRefs.Add(new CellAddress(comment.@ref), comment);
                }
            }
        }
    }

}



