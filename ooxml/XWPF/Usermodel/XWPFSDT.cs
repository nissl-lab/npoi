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
using System;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.Collections.Generic;
namespace NPOI.XWPF.UserModel
{

    /**
     * Experimental class to offer rudimentary Read-only Processing of 
     *  of StructuredDocumentTags/ContentControl
     *  
     *
     *
     * WARNING - APIs expected to change rapidly
     * 
     */
    public class XWPFSDT : IBodyElement, IRunBody, ISDTContents, IRunElement
    {
        private String title;
        private String tag;
        private XWPFSDTContent content;
        private IBody part;

        public XWPFSDT(CT_SdtRun sdtRun, IBody part)
        {
            this.part = part;
            this.content = new XWPFSDTContent(sdtRun.sdtContent, part, this);
            CT_SdtPr pr = sdtRun.sdtPr;
            List<CT_String> aliases = pr.GetObjectList<CT_String>(SdtPrElementType.alias);
            if (aliases != null && aliases.Count > 0)
            {
                title = aliases[0].val;
            }
            else
            {
                title = "";
            }
            CT_String[] array = pr.GetObjectList<CT_String>(SdtPrElementType.tag).ToArray();
            if (array != null && array.Length > 0)
            {
                tag = array[0].val;
            }
            else
            {
                tag = "";
            }

        }
        public XWPFSDT(CT_SdtBlock block, IBody part)
        {
            this.part = part;
            this.content = new XWPFSDTContent(block.sdtContent, part, this);
            CT_SdtPr pr = block.sdtPr;
            List<CT_String> aliases = pr.GetObjectList<CT_String>(SdtPrElementType.alias);
            if (aliases != null && aliases.Count > 0)
            {
                title = aliases[0].val;
            }
            else
            {
                title = "";
            }
            CT_String[] array = pr.GetObjectList<CT_String>(SdtPrElementType.tag).ToArray();
            if (array != null && array.Length > 0)
            {
                tag = array[0].val;
            }
            else
            {
                tag = "";
            }

        }
        public String Title
        {
            get
            {
                return title;
            }
        }
        public String Tag
        {
            get
            {
                return tag;
            }
        }
        public XWPFSDTContent Content
        {
            get
            {
                return content;
            }
        }

        public IBody Body
        {
            get
            {
                // TODO Auto-generated method stub
                return null;
            }
        }

        public POIXMLDocumentPart Part
        {
            get
            {
                return part.Part;
            }
        }

        public BodyType PartType
        {
            get
            {
                return BodyType.CONTENTCONTROL;
            }
        }

        public BodyElementType ElementType
        {
            get
            {
                return BodyElementType.CONTENTCONTROL;
            }
        }

        public XWPFDocument Document
        {
            get
            {
                return part.GetXWPFDocument();
            }
        }
    }

}