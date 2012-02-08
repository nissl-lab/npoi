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
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.XSSF.Model;
using System;
using NPOI.SS.Util;
using NPOI.OpenXmlFormats.Dml;
namespace NPOI.XSSF.UserModel
{



    public class XSSFComment : IComment
    {

        private CT_Comment _comment;
        private CommentsTable _comments;
        private CT_Shape _vmlShape;

        /**
         * cached reference to the string with the comment text
         */
        private XSSFRichTextString _str;

        /**
         * Creates a new XSSFComment, associated with a given
         *  low level comment object.
         */
        public XSSFComment(CommentsTable comments, CT_Comment comment, CT_Shape vmlShape)
        {
            _comment = comment;
            _comments = comments;
            _vmlShape = vmlShape;
        }

        /**
         *
         * @return Name of the original comment author. Default value is blank.
         */
        public String Author
        {
            get
            {
                return _comments.GetAuthor((int)_comment.authorId);
            }
        }

        /**
         * Name of the original comment author. Default value is blank.
         *
         * @param author the name of the original author of the comment
         */
        public void SetAuthor(String author)
        {
            _comment.authorId = (
                    (uint)_comments.FindAuthor(author)
            );
        }

        /**
         * @return the 0-based column of the cell that the comment is associated with.
         */
        public int Column
        {
            get
            {
                return new CellReference(_comment.@ref).Col;
            }
        }

        /**
         * @return the 0-based row index of the cell that the comment is associated with.
         */
        public int Row
        {
            get
            {
                return new CellReference(_comment.@ref).Row;
            }
        }

        /**
         * @return whether the comment is visible
         */
        public bool IsVisible()
        {
            bool visible = false;
            if (_vmlShape != null)
            {
                String style = _vmlShape.GetStyle();
                visible = style != null && style.IndexOf("visibility:visible") != -1;
            }
            return visible;
        }

        /**
         * @param visible whether the comment is visible
         */
        public void SetVisible(bool visible)
        {
            if (_vmlShape != null)
            {
                String style;
                if (visible) style = "position:absolute;visibility:visible";
                else style = "position:absolute;visibility:hidden";
                _vmlShape.SetStyle(style);
            }
        }

        /**
         * Set the column of the cell that Contains the comment
         *
         * @param col the 0-based column of the cell that Contains the comment
         */
        public void SetColumn(int col)
        {
            String oldRef = _comment.@ref;

            CellReference ref1 = new CellReference(Row, col);
            _comment.@ref = (ref1.FormatAsString());
            _comments.ReferenceUpdated(oldRef, _comment);

            if (_vmlShape != null)
            {
                _vmlShape.GetClientDataArray(0).SetColumnArray(
                      new BigInt[] { new Bigint(String.ValueOf(col)) }
                );

                // There is a very odd xmlbeans bug when changing the column
                //  arrays which can lead to corrupt pointer
                // This call seems to fix them again... See bug #50795
                _vmlShape.GetClientDataList().ToString();
            }
        }

        /**
         * Set the row of the cell that Contains the comment
         *
         * @param row the 0-based row of the cell that Contains the comment
         */
        public void SetRow(int row)
        {
            String oldRef = _comment.@ref;

            String newRef =
                (new CellReference(row, Column)).FormatAsString();
            _comment.@ref = (newRef);
            _comments.ReferenceUpdated(oldRef, _comment);

            if (_vmlShape != null) _vmlShape.GetClientDataArray(0).SetRowArray(0, new Bigint(String.ValueOf(row)));
        }

        /**
         * @return the rich text string of the comment
         */
        public XSSFRichTextString GetString()
        {
            if (_str == null)
            {
                CT_Rst rst = _comment.text;
                if (rst != null) _str = new XSSFRichTextString(_comment.GetText());
            }
            return _str;
        }

        /**
         * Sets the rich text string used by this comment.
         *
         * @param string  the XSSFRichTextString used by this object.
         */
        public void SetString(IRichTextString str)
        {
            if (!(str is XSSFRichTextString))
            {
                throw new ArgumentException("Only XSSFRichTextString argument is supported");
            }
            _str = (XSSFRichTextString)str;
            _comment.text = (_str.GetCTRst());
        }

        public void SetString(String str)
        {
            SetString(new XSSFRichTextString(str));
        }

        /**
         * @return the xml bean holding this comment's properties
         */
        protected CT_Comment GetCTComment()
        {
            return _comment;
        }

        protected CT_Shape GetCTShape()
        {
            return _vmlShape;
        }
    }
}

