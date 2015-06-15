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
     * A run of text with a Hyperlink applied to it.
     * Any given Hyperlink may be made up of multiple of these.
     */
    public class XWPFHyperlinkRun : XWPFRun
    {
        private CT_Hyperlink1 hyperlink;

        public XWPFHyperlinkRun(CT_Hyperlink1 hyperlink, CT_R Run, IRunBody p)
            : base(Run, p)
        {
            this.hyperlink = hyperlink;
        }

        public CT_Hyperlink1 GetCTHyperlink()
        {
            return hyperlink;
        }

        public String GetAnchor()
        {
            return hyperlink.anchor;
        }

        /**
         * Returns the ID of the hyperlink, if one is Set.
         */
        public String GetHyperlinkId()
        {
            return hyperlink.id;
        }
        public void SetHyperlinkId(String id)
        {
            hyperlink.id = (id);
        }

        /**
         * If this Hyperlink is an external reference hyperlink,
         *  return the object for it.
         */
        public XWPFHyperlink GetHyperlink(XWPFDocument document)
        {
            String id = GetHyperlinkId();
            if (id == null)
                return null;

            return document.GetHyperlinkByID(id);
        }
    }

}