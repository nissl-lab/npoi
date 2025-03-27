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

namespace NPOI.XDDF.UserModel
{
    using NPOI.OpenXmlFormats.Dml;
    public class XDDFLineProperties
    {
        private CT_LineProperties props;

        public XDDFLineProperties()
            : this(new CT_LineProperties())
        {

        }
        public XDDFLineProperties(CT_LineProperties properties)
        {
            this.props = properties;
        }
        public CT_LineProperties GetXmlObject()
        {
            return props;
        }

        public PenAlignment? PenAlignment
        {
            get
            {
                if(props.algnSpecified)
                {
                    return PenAlignmentExtensions.ValueOf(props.algn);
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
                    props.algnSpecified = false;
                }
                else
                {
                    props.algn = value.Value.ToST_PenAlignment();
                }
            }
        }

        public LineCap? LineCap
        {
            get
            {
                if(props.capSpecified)
                {
                    return LineCapExtensions.ValueOf(props.cap);
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
                    props.algnSpecified = false;
                }
                else
                {
                    props.cap = value.Value.ToST_LineCap();
                }
            }
        }

        public CompoundLine? CompoundLine
        {
            get
            {
                if(props.cmpdSpecified)
                {
                    return CompoundLineExtensions.ValueOf(props.cmpd);
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
                    props.cmpdSpecified = false;
                }
                else
                {
                    props.cmpd = value.Value.ToST_CompoundLine();
                }
            }
        }

        public XDDFDashStop AddDashStop()
        {
            if(!props.IsSetCustDash())
            {
                props.AddNewCustDash();
            }
            return new XDDFDashStop(props.custDash.AddNewDs());
        }

        public XDDFDashStop InsertDashStop(int index)
        {
            if(!props.IsSetCustDash())
            {
                props.AddNewCustDash();
            }
            return new XDDFDashStop(props.custDash.InsertNewDs(index));
        }

        public void RemoveDashStop(int index)
        {
            if(props.IsSetCustDash())
            {
                props.custDash.RemoveDs(index);
            }
        }

        public XDDFDashStop GetDashStop(int index)
        {
            if(props.IsSetCustDash())
            {
                return new XDDFDashStop(props.custDash.GetDsArray(index));
            }
            else
            {
                return null;
            }
        }

        public List<XDDFDashStop> GetDashStops()
        {
            if(props.IsSetCustDash())
            {
                return Collections.unmodifiableList(props
                    .CustDash
                    .DsList
                    .stream()
                    .map(ds-> new XDDFDashStop(ds))
                    .collect(Collectors.ToList()));
            }
            else
            {
                return new List<XDDFDashStop>();
            }
        }

        public int CountDashStops()
        {
            if(props.IsSetCustDash())
            {
                return props.custDash.ds.Count;
            }
            else
            {
                return 0;
            }
        }

        public XDDFPresetLineDash GetPresetDash()
        {
            if(props.IsSetPrstDash())
            {
                return new XDDFPresetLineDash(props.prstDash);
            }
            else
            {
                return null;
            }
        }

        public void SetPresetDash(XDDFPresetLineDash properties)
        {
            if(properties == null)
            {
                if(props.IsSetPrstDash())
                {
                    props.UnsetPrstDash();
                }
            }
            else
            {
                props.prstDash = properties.GetXmlObject();
            }
        }

        public XDDFExtensionList GetExtensionList()
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

        public void SetExtensionList(XDDFExtensionList list)
        {
            if(list == null)
            {
                if(props.IsSetExtLst())
                {
                    props.UnsetExtLst();
                }
            }
            else
            {
                props.extLst = list.GetXmlObject();
            }
        }

        public IXDDFFillProperties GetFillProperties()
        {
            if(props.IsSetGradFill())
            {
                return new XDDFGradientFillProperties(props.gradFill);
            }
            else if(props.IsSetNoFill())
            {
                return new XDDFNoFillProperties(props.noFill);
            }
            else if(props.IsSetPattFill())
            {
                return new XDDFPatternFillProperties(props.pattFill);
            }
            else if(props.IsSetSolidFill())
            {
                return new XDDFSolidFillProperties(props.solidFill);
            }
            else
            {
                return null;
            }
        }

        public void SetFillProperties(IXDDFFillProperties properties)
        {
            if(props.IsSetGradFill())
            {
                props.UnsetGradFill();
            }
            if(props.IsSetNoFill())
            {
                props.UnsetNoFill();
            }
            if(props.IsSetPattFill())
            {
                props.UnsetPattFill();
            }
            if(props.IsSetSolidFill())
            {
                props.UnsetSolidFill();
            }
            if(properties == null)
            {
                return;
            }
            if(properties is XDDFGradientFillProperties)
            {
                props.gradFill = ((XDDFGradientFillProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFNoFillProperties)
            {
                props.noFill = ((XDDFNoFillProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFPatternFillProperties)
            {
                props.pattFill = ((XDDFPatternFillProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFSolidFillProperties)
            {
                props.solidFill = ((XDDFSolidFillProperties) properties).GetXmlObject();
            }
        }

        public IXDDFLineJoinProperties GetLineJoinProperties()
        {
            if(props.IsSetBevel())
            {
                return new XDDFLineJoinBevelProperties(props.bevel);
            }
            else if(props.IsSetMiter())
            {
                return new XDDFLineJoinMiterProperties(props.miter);
            }
            else if(props.IsSetRound())
            {
                return new XDDFLineJoinRoundProperties(props.round);
            }
            else
            {
                return null;
            }
        }

        public void SetLineJoinProperties(IXDDFLineJoinProperties properties)
        {
            if(props.IsSetBevel())
            {
                props.UnsetBevel();
            }
            if(props.IsSetMiter())
            {
                props.UnsetMiter();
            }
            if(props.IsSetRound())
            {
                props.UnsetRound();
            }
            if(properties == null)
            {
                return;
            }
            if(properties is XDDFLineJoinBevelProperties)
            {
                props.bevel = ((XDDFLineJoinBevelProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFLineJoinMiterProperties)
            {
                props.miter = ((XDDFLineJoinMiterProperties) properties).GetXmlObject();
            }
            else if(properties is XDDFLineJoinRoundProperties)
            {
                props.round = ((XDDFLineJoinRoundProperties) properties).GetXmlObject();
            }
        }

        public XDDFLineEndProperties HeadEnd
        {
            get
            {
                if(props.IsSetHeadEnd())
                {
                    return new XDDFLineEndProperties(props.headEnd);
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
                    if(props.IsSetHeadEnd())
                    {
                        props.UnsetHeadEnd();
                    }
                }
                else
                {
                    props.headEnd = value.GetXmlObject();
                }
            }
        }

        public XDDFLineEndProperties TailEnd
        {
            get
            {
                if(props.IsSetTailEnd())
                {
                    return new XDDFLineEndProperties(props.tailEnd);
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
                    if(props.IsSetTailEnd())
                    {
                        props.UnsetTailEnd();
                    }
                }
                else
                {
                    props.tailEnd = value.GetXmlObject();
                }
            }
        }

        public int? Width
        {
            get
            {
                if(props.wSpecified)
                {
                    return props.w;
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
                    props.wSpecified = false;
                }
                else
                {
                    props.w = value.Value;
                }
            }
        }
    }
}


