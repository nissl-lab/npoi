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
    using System.Text;
    using NPOI.OpenXmlFormats.Wordprocessing;

    /**
     * Sketch of XWPF comment class
     * 
    * @author Yury Batrakov (batrakov at gmail.com)
     * 
     */
    public class XWPFComment
    {
        protected String id;
        protected String author;
        protected StringBuilder text;

        public XWPFComment(CT_Comment comment, XWPFDocument document)
        {
            text = new StringBuilder();
            id = comment.id.ToString();
            author = comment.author;

            foreach(CT_P ctp in comment.GetPList())
            {
                XWPFParagraph p = new XWPFParagraph(ctp, document);
                text.Append(p.Text);
            }
        }

        public String Id
        {
            get
            {
                return id;
            }
        }

        public String Author
        {
            get
            {
                return author;
            }
        }

        public String Text
        {
            get
            {
                return text.ToString();
            }
        }
    }

}