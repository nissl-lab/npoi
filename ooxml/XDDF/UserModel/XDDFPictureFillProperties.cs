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
    using NPOI.Util;


    using NPOI.OpenXmlFormats.Dml;
    public class XDDFPictureFillProperties : IXDDFFillProperties
    {
        private CT_BlipFillProperties props;

        public XDDFPictureFillProperties()
            : this(new CT_BlipFillProperties())
        {
        }

        public XDDFPictureFillProperties(CT_BlipFillProperties properties)
        {
            this.props = properties;
        }
        public CT_BlipFillProperties GetXmlObject()
        {
            return props;
        }

        public XDDFPicture Picture
        {
            get
            {
                if(props.IsSetBlip())
                {
                    return new XDDFPicture(props.blip);
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
                    props.UnsetBlip();
                }
                else
                {
                    props.blip = value.GetXmlObject();
                }
            }
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

        public long? Dpi
        {
            get
            {
                if(props.dpiSpecified)
                {
                    return props.dpi;
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
                    props.dpiSpecified = false;
                }
                else
                {
                    props.dpi = (uint)value.Value;
                }
            }
        }

        public XDDFRelativeRectangle SourceRectangle
        {
            get
            {
                if(props.IsSetSrcRect())
                {
                    return new XDDFRelativeRectangle(props.srcRect);
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
                    if(props.IsSetSrcRect())
                    {
                        props.UnsetSrcRect();
                    }
                }
                else
                {
                    props.srcRect = value.GetXmlObject();
                }
            }
        }

        public XDDFStretchInfoProperties StretchInfoProperties
        {
            get
            {
                if(props.IsSetStretch())
                {
                    return new XDDFStretchInfoProperties(props.stretch);
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
                    if(props.IsSetStretch())
                    {
                        props.UnsetStretch();
                    }
                }
                else
                {
                    props.stretch = value.GetXmlObject();
                }
            }
        }

        public XDDFTileInfoProperties TileInfoProperties
        {
            get
            {
                if(props.IsSetTile())
                {
                    return new XDDFTileInfoProperties(props.tile);
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
                    if(props.IsSetTile())
                    {
                        props.UnsetTile();
                    }
                }
                else
                {
                    props.tile = value.GetXmlObject();
                }
            }
        }
    }
}
