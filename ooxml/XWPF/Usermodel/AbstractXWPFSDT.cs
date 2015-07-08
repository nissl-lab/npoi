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
     * Experimental abstract class that is a base for XWPFSDT and XWPFSDTCell
     * <p/>
     * WARNING - APIs expected to change rapidly.
     * <p/>
     * These classes have so far been built only for Read-only Processing.
     */
    public abstract class AbstractXWPFSDT : ISDTContents
    {
        private String title;
        private String tag;
        private IBody part;

        public AbstractXWPFSDT(CT_SdtPr pr, IBody part)
        {

            CT_String[] aliases = pr.GetAliasArray();
            if (aliases != null && aliases.Length > 0)
            {
                title = aliases[0].val;
            }
            else
            {
                title = "";
            }
            CT_String[] tags = pr.GetAliasArray();
            if (tags != null && tags.Length > 0)
            {
                tag = tags[0].val;
            }
            else
            {
                tag = "";
            }
            this.part = part;

        }

        /**
         * @return first SDT Title
         */
        public String GetTitle()
        {
            return title;
        }

        /**
         * @return first SDT Tag
         */
        public String GetTag()
        {
            return tag;
        }

        /**
         * @return the content object
         */
        public abstract ISDTContent Content { get; }

        /**
         * @return null
         */
        public IBody GetBody()
        {
            return null;
        }

        /**
         * @return document part
         */
        public POIXMLDocumentPart GetPart()
        {
            return part.Part;
        }

        /**
         * @return partType
         */
        public BodyType GetPartType()
        {
            return BodyType.CONTENTCONTROL;
        }

        /**
         * @return element type
         */
        public BodyElementType GetElementType()
        {
            return BodyElementType.CONTENTCONTROL;
        }

        public XWPFDocument GetDocument()
        {
            return part.GetXWPFDocument();
        }
    }

}