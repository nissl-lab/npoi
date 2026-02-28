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
    using NPOI.OpenXmlFormats.Wordprocessing;

    /**
     * Default Paragraph style, from which other styles will override
     * TODO Share logic with {@link XWPFParagraph} which also uses CTPPr
     */
    public class XWPFDefaultParagraphStyle
    {
        private CT_PPr ppr;

        public XWPFDefaultParagraphStyle(CT_PPr ppr)
        {
            this.ppr = ppr;
        }

        protected internal CT_PPr GetPPr()
        {
            return ppr;
        }

        public int SpacingAfter
        {
            get
            {
                if (ppr.IsSetSpacing())
                    return (int)ppr.spacing.after;
                return -1;
            }
        }
    }

}