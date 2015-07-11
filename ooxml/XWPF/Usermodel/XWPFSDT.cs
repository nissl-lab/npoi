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
    public class XWPFSDT : AbstractXWPFSDT, IBodyElement, IRunBody, ISDTContents, IRunElement
    {
        private ISDTContent content;

        public XWPFSDT(CT_SdtRun sdtRun, IBody part)
            : base(sdtRun.sdtPr, part)
        {
            this.content = new XWPFSDTContent(sdtRun.sdtContent, part, this);

        }
        public XWPFSDT(CT_SdtBlock block, IBody part)
            : base(block.sdtPr, part)
        {
            this.content = new XWPFSDTContent(block.sdtContent, part, this);
        }

        public override ISDTContent Content
        {
            get
            {
                return content;
            }
        }

        public XWPFDocument Document
        {
            get { return GetDocument(); }
        }

        public POIXMLDocumentPart Part
        {
            get { return GetPart(); }
        }

        public IBody Body
        {
            get { return GetBody(); }
        }

        public BodyType PartType
        {
            get { return GetPartType(); }
        }

        public BodyElementType ElementType
        {
            get { return GetElementType(); }
        }
    }

}