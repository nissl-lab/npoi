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
using System.Linq;

namespace NPOI.XDDF.UserModel
{
    using NPOI.OpenXmlFormats.Dml;
    public class XDDFGradientFillProperties : IXDDFFillProperties
    {
        private CT_GradientFillProperties props;

        public XDDFGradientFillProperties()
            : this(new CT_GradientFillProperties())
        {

        }

        public XDDFGradientFillProperties(CT_GradientFillProperties properties)
        {
            this.props = properties;
        }
        public CT_GradientFillProperties GetXmlObject()
        {
            return props;
        }

        public bool? IsRotatingWithShape
        {
            get
            {
                if(props.rotWithShapeSpecified)
                {
                    return props.rotWithShape;
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
                    props.rotWithShapeSpecified = false;
                }
                else
                {
                    props.rotWithShape = value.Value;
                }
            }
        }

        public TileFlipMode? TileFlipMode
        {
            get
            {
                if(props.flipSpecified)
                {
                    return TileFlipModeExtensions.ValueOf(props.flip);
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
                    props.flipSpecified = false;
                }
                else
                {
                    props.flip = value.Value.ToST_TileFlipMode();
                }
            }
        }

        public XDDFGradientStop AddGradientStop()
        {
            if(!props.IsSetGsLst())
            {
                props.AddNewGsLst();
            }
            return new XDDFGradientStop(props.gsLst.AddNewGs());
        }

        public XDDFGradientStop InsertGradientStop(int index)
        {
            if(!props.IsSetGsLst())
            {
                props.AddNewGsLst();
            }
            return new XDDFGradientStop(props.gsLst.InsertNewGs(index));
        }

        public void RemoveGradientStop(int index)
        {
            if(props.IsSetGsLst())
            {
                props.gsLst.RemoveGs(index);
            }
        }

        public XDDFGradientStop GetGradientStop(int index)
        {
            if(props.IsSetGsLst())
            {
                return new XDDFGradientStop(props.gsLst.GetGsArray(index));
            }
            else
            {
                return null;
            }
        }

        public List<XDDFGradientStop> GetGradientStops()
        {
            if(props.IsSetGsLst())
            {
                return props.gsLst.gs.Select(x => new XDDFGradientStop(x)).ToList();
            }
            else
            {
                return new List<XDDFGradientStop>();
            }
        }

        public int CountGradientStops()
        {
            if(props.IsSetGsLst())
            {
                return props.gsLst.SizeOfGsArray();
            }
            else
            {
                return 0;
            }
        }

        public XDDFLinearShadeProperties LinearShadeProperties
        {
            get
            {
                if(props.IsSetLin())
                {
                    return new XDDFLinearShadeProperties(props.lin);
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
                    if(props.IsSetLin())
                    {
                        props.UnsetLin();
                    }
                }
                else
                {
                    props.lin = value.GetXmlObject();
                }
            }
        }

        public XDDFPathShadeProperties PathShadeProperties
        {
            get
            {
                if(props.IsSetPath())
                {
                    return new XDDFPathShadeProperties(props.path);
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
                    if(props.IsSetPath())
                    {
                        props.UnsetPath();
                    }
                }
                else
                {
                    props.path = value.GetXmlObject();
                }
            }
        }

        public XDDFRelativeRectangle TileRectangle
        {
            get
            {
                if(props.IsSetTileRect())
                {
                    return new XDDFRelativeRectangle(props.tileRect);
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
                    if(props.IsSetTileRect())
                    {
                        props.UnsetTileRect();
                    }
                }
                else
                {
                    props.tileRect = value.GetXmlObject();
                }
            }
        }
    }
}