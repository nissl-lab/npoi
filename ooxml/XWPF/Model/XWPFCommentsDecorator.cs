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
namespace NPOI.XWPF.Model
{
    using System;

    using NPOI.XWPF.UserModel;
    using System.Text;
    using NPOI.OpenXmlFormats.Wordprocessing;

    /**
     * Decorator class for XWPFParagraph allowing to add comments 
     * found in paragraph to its text
     *
     * @author Yury Batrakov (batrakov at gmail.com)
     * 
     */
    public class XWPFCommentsDecorator : XWPFParagraphDecorator
    {
        private StringBuilder commentText;

        public XWPFCommentsDecorator(XWPFParagraphDecorator nextDecorator):
            this(nextDecorator.paragraph, nextDecorator)
        {
        }
        public XWPFCommentsDecorator(XWPFParagraph paragraph, XWPFParagraphDecorator nextDecorator)
            : base(paragraph, nextDecorator)
        {
            ;

            XWPFComment comment;
            commentText = new StringBuilder();

            foreach (CT_MarkupRange anchor in paragraph.GetCTP().GetCommentRangeStartList())
            {
                if ((comment = paragraph.Document.GetCommentByID(anchor.id)) != null)
                    commentText.Append("\tComment by " + comment.Author + ": " + comment.Text);
            }
        }

        public String GetCommentText()
        {
            return commentText.ToString();
        }

        public override String Text
        {
            get
            {
                return base.Text + commentText;
            }
        }
    }
}
