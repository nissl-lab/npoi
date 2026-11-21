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

namespace NPOI.XDDF.UserModel.Chart
{
    using NPOI.OpenXmlFormats.Dml.Chart;
    public class XDDFLayout
    {
        private CT_Layout layout;

        public XDDFLayout()
                : this(new CT_Layout())
        {
        }
        internal XDDFLayout(CT_Layout layout)
        {
            this.layout = layout;
        }
        internal CT_Layout GetXmlObject()
        {
            return layout;
        }

        public XDDFChartExtensionList ExtensionList
        {
            get
            {
                if(layout.IsSetExtLst())
                {
                    return new XDDFChartExtensionList(layout.extLst);
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
                    if(layout.IsSetExtLst())
                    {
                        layout.UnsetExtLst();
                    }
                }
                else
                {
                    layout.extLst = value.GetXmlObject();
                }
            }
        }

        public XDDFManualLayout ManualLayout
        {
            get
            {
                if(layout.IsSetManualLayout())
                {
                    return new XDDFManualLayout(layout);
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
                    if(layout.IsSetManualLayout())
                    {
                        layout.UnsetManualLayout();
                    }
                }
                else
                {
                    layout.manualLayout = value.GetXmlObject();
                }
            }
        }
    }
}
