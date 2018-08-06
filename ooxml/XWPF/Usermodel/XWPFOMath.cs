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
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using NPOI.OpenXmlFormats.Shared;
    using NPOI.XWPF.Usermodel;

    public class XWPFOMath : IRunBody
    {
        protected CT_OMath oMath;
        protected IRunBody parent;

        protected List<XWPFSharedRun> runs;
        protected XWPFDocument document;

        public XWPFOMath(CT_OMath oMath, IRunBody p)
        {
            this.oMath = oMath;
            this.parent = p;
            this.document = p.Document;

            runs = new List<XWPFSharedRun>();

            BuildRunsInOrderFromXml(oMath.Items);
        }

        public XWPFDocument Document
        {
            get { return document; }
        }

        public POIXMLDocumentPart Part { get { return parent.Part; } }

        private void BuildRunsInOrderFromXml(ArrayList items)
        {
            foreach (object o in items)
            {
                if (o is CT_R)
                {
                    runs.Add(new XWPFSharedRun(o as CT_R, this));
                }
            }
        }

        public XWPFSharedRun CreateRun()
        {
            XWPFSharedRun run = new XWPFSharedRun(oMath.AddNewR(), this);
            runs.Add(run);
            return run;
        }

        public IList<XWPFSharedRun> Runs
        {
            get
            {
                return runs.AsReadOnly();
            }
        }


    }
}
