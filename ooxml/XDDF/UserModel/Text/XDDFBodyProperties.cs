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
    public class XDDFBodyProperties
    {
        private CT_TextBodyProperties props;
        internal XDDFBodyProperties(CT_TextBodyProperties properties)
        {
            this.props = properties;
        }
        internal CT_TextBodyProperties GetXmlObject()
        {
            return props;
        }

        public AnchorType? Anchoring
        {
            get
            {
                if(props.IsSetAnchor())
                {
                    return AnchorTypeExtensions.ValueOf(props.anchor);
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
                    if(props.IsSetAnchor())
                    {
                        props.UnsetAnchor();
                    }
                }
                else
                {
                    props.anchor = value.Value.ToST_TextAnchoringType();
                }
            }
        }

        public bool? IsAnchorCentered
        {
            get
            {
                if(props.anchorCtrSpecified)
                {
                    return props.anchorCtr;
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
                    props.anchorCtrSpecified = false;
                }
                else
                {
                    props.anchorCtr = value.Value;
                }
            }
        }

        public IXDDFAutoFit AutoFit
        {
            get
            {
                if(props.IsSetNoAutofit())
                {
                    return new XDDFNoAutoFit(props.noAutofit);
                }
                else if(props.IsSetNormAutofit())
                {
                    return new XDDFNormalAutoFit(props.normAutofit);
                }
                else if(props.IsSetSpAutoFit())
                {
                    return new XDDFShapeAutoFit(props.spAutoFit);
                }
                return new XDDFNormalAutoFit();
            }
            set
            {
                if(props.IsSetNoAutofit())
                {
                    props.UnsetNoAutofit();
                }
                if(props.IsSetNormAutofit())
                {
                    props.UnsetNormAutofit();
                }
                if(props.IsSetSpAutoFit())
                {
                    props.UnsetSpAutoFit();
                }
                IXDDFAutoFit autofit = value;
                if(autofit is XDDFNoAutoFit)
                {
                    props.noAutofit = ((XDDFNoAutoFit) autofit).GetXmlObject();
                }
                else if(autofit is XDDFNormalAutoFit)
                {
                    props.normAutofit = ((XDDFNormalAutoFit) autofit).GetXmlObject();
                }
                else if(autofit is XDDFShapeAutoFit)
                {
                    props.spAutoFit = ((XDDFShapeAutoFit) autofit).GetXmlObject();
                }
            }
        }


        public XDDFExtensionList ExtensionList
        {
            get
            {
                if(props.IsSetExtLst())
                {
                    return new XDDFExtensionList(props.extLst);
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
                    if(props.IsSetExtLst())
                    {
                        props.UnsetExtLst();
                    }
                }
                else
                {
                    props.extLst = value.GetXmlObject();
                }
            }
        }

        public double? BottomInset
        {
            get
            {
                if(props.IsSetBIns())
                {
                    return Units.ToPoints(props.bIns);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null || Double.IsNaN(value.Value))
                {
                    if(props.IsSetBIns())
                    {
                        props.UnsetBIns();
                    }
                }
                else
                {
                    props.bIns = Units.ToEMU(value.Value);
                }
            }
        }

        public double? LeftInset
        {
            get
            {
                if(props.IsSetLIns())
                {
                    return Units.ToPoints(props.lIns);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null || Double.IsNaN(value.Value))
                {
                    if(props.IsSetLIns())
                    {
                        props.UnsetLIns();
                    }
                }
                else
                {
                    props.lIns = Units.ToEMU(value.Value);
                }
            }
        }

        public double? RightInset
        {
            get
            {
                if(props.IsSetRIns())
                {
                    return Units.ToPoints(props.rIns);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null || Double.IsNaN(value.Value))
                {
                    if(props.IsSetRIns())
                    {
                        props.UnsetRIns();
                    }
                }
                else
                {
                    props.rIns = Units.ToEMU(value.Value);
                }
            }
        }

        public double? TopInset
        {
            get
            {
                if(props.IsSetTIns())
                {
                    return Units.ToPoints(props.tIns);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value == null || Double.IsNaN(value.Value))
                {
                    if(props.IsSetTIns())
                    {
                        props.UnsetTIns();
                    }
                }
                else
                {
                    props.tIns = Units.ToEMU(value.Value);
                }
            }
        }

        public bool? HasParagraphSpacing
        {
            get
            {
                if(props.spcFirstLastParaSpecified)
                {
                    return props.spcFirstLastPara;
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
                    props.spcFirstLastPara = false;
                }
                else
                {
                    props.spcFirstLastPara = value.Value;
                }
            }
        }

        public bool? RightToLeft
        {
            get
            {
                if(props.rtlCol)
                {
                    return props.rtlCol;
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
                    props.rtlColSpecified = false;
                }
                else
                {
                    props.rtlCol = value.Value;
                }
            }
        }
    }
}
