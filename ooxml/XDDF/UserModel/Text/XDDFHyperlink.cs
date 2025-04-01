/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel.Text
{
    using NPOI.Util;


    using NPOI.XDDF.UserModel;

    using NPOI.OpenXmlFormats.Dml;
    using Org.BouncyCastle.Utilities.Collections;

    public class XDDFHyperlink
    {
        private CT_Hyperlink link;

        public XDDFHyperlink(string id)
            : this(new CT_Hyperlink())
        {
            this.link.id = id;
        }

        public XDDFHyperlink(string id, string action)
            : this(id)
        {

            this.link.action = action;
        }
        protected XDDFHyperlink(CT_Hyperlink link)
        {
            this.link = link;
        }
        internal CT_Hyperlink GetXmlObject()
        {
            return link;
        }

        public string Action
        {
            get
            {
                if(link.actionSpecified)
                {
                    return link.action;
                }
                else
                {
                    return null;
                }
            }
        }

        public string Id
        {
            get
            {
                if(link.idSpecified)
                {
                    return link.id;
                }
                else
                {
                    return null;
                }
            }
        }

        public bool? EndSound
        {
            get
            {
                if(link.endSndSpecified)
                {
                    return link.endSnd;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    link.endSndSpecified = false;
                }
                else
                {
                    link.endSnd = value.Value;
                }
            }
        }

        public bool? HighlightClick
        {
            get
            {
                if(link.highlightClickSpecified)
                {
                    return link.highlightClick;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    link.highlightClickSpecified = false;
                }
                else
                {
                    link.highlightClick = value.Value;
                }
            }
        }


        public bool? History
        {
            get
            {
                if(link.historySpecified)
                {
                    return link.history;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    link.historySpecified = false;
                }
                else
                {
                    link.history = value.Value;
                }
            }
        }

        public string InvalidURL
        {
            get
            {
                if(link.invalidUrlSpecified)
                {
                    return link.invalidUrl;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    link.invalidUrlSpecified = false;
                }
                else
                {
                    link.invalidUrl = value;
                }
            }
        }


        public string TargetFrame
        {
            get
            {
                if(link.tgtFrameSpecified)
                {
                    return link.tgtFrame;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    link.tgtFrameSpecified = false;
                }
                else
                {
                    link.tgtFrame = value;
                }
            }
        }


        public string Tooltip
        {
            get
            {
                if(link.tooltipSpecified)
                {
                    return link.tooltip;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    link.tooltipSpecified = false;
                }
                else
                {
                    link.tooltip = value;
                }
            }
        }

        public XDDFExtensionList ExtensionList
        {
            get
            {
                if(link.IsSetExtLst())
                {
                    return new XDDFExtensionList(link.extLst);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null)
                {
                    if(link.IsSetExtLst())
                    {
                        link.UnsetExtLst();
                    }
                }
                else
                {
                    link.extLst = value.GetXmlObject();
                }
            }
        }
    }
}
