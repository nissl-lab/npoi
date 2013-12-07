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
     * Decorator class for XWPFParagraph allowing to add hyperlinks 
     *  found in paragraph to its text.
     *  
     * Note - Adds the hyperlink at the end, not in the right place...
     *  
     * @deprecated Use {@link XWPFHyperlinkRun} instead
     */
    public class XWPFHyperlinkDecorator : XWPFParagraphDecorator
    {
        private StringBuilder hyperlinkText;

        /**
         * @param nextDecorator The next decorator to use
         * @param outputHyperlinkUrls Should we output the links too, or just the link text?
         */
        public XWPFHyperlinkDecorator(XWPFParagraphDecorator nextDecorator, bool outputHyperlinkUrls) :
            this(nextDecorator.paragraph, nextDecorator, outputHyperlinkUrls)
        {
        }

        /**
         * @param prgrph The paragraph of text to work on
         * @param outputHyperlinkUrls Should we output the links too, or just the link text?
         */
        public XWPFHyperlinkDecorator(XWPFParagraph prgrph, XWPFParagraphDecorator nextDecorator, bool outputHyperlinkUrls)
            : base(prgrph, nextDecorator)
        {
            hyperlinkText = new StringBuilder();

            // loop over hyperlink anchors
            foreach (CT_Hyperlink1 link in paragraph.GetCTP().GetHyperlinkList())
            {
                foreach (CT_R r in link.GetRList())
                {
                    // Loop over text Runs
                    foreach (CT_Text text in r.GetTList())
                    {
                        hyperlinkText.Append(text.Value);
                    }
                }
                if (outputHyperlinkUrls && paragraph.Document.GetHyperlinkByID(link.id) != null)
                {
                    hyperlinkText.Append(" <" + paragraph.Document.GetHyperlinkByID(link.id).URL + ">");
                }
            }
        }

        public override String Text
        {
            get
            {
                return base.Text + hyperlinkText;
            }
        }
    }

}