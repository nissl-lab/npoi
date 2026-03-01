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


    public class XDDFTransform2D
    {
        private CT_Transform2D transform;

        public XDDFTransform2D(CT_Transform2D transform)
        {
            this.transform = transform;
        }
        public CT_Transform2D GetXmlObject()
        {
            return transform;
        }

        public bool? FlipHorizontal
        {
            get
            {
                if(transform.flipHSpecified)
                {
                    return transform.flipH;
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
                    transform.flipHSpecified = true;
                }
                else
                {
                    transform.flipH = value.Value;
                }
            }
        }

        public bool? GetFlipVertical
        {
            get
            {
                if(transform.flipVSpecified)
                {
                    return transform.flipV;
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
                    transform.flipVSpecified = true;
                }
                else
                {
                    transform.flipV = value.Value;
                }
            }
        }

        public XDDFPositiveSize2D Extension
        {
            get
            {
                if(transform.IsSetExt())
                {
                    return new XDDFPositiveSize2D(transform.ext);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                CT_PositiveSize2D xformExt;
                if(value == null)
                {
                    if(transform.IsSetExt())
                    {
                        transform.UnsetExt();
                    }
                    return;
                }
                else if(transform.IsSetExt())
                {
                    xformExt = transform.ext;
                }
                else
                {
                    xformExt = transform.AddNewExt();
                }
                xformExt.cx = value.X;
                xformExt.cy = value.Y;
            }
        }

        public XDDFPoint2D Offset
        {
            get
            {
                if(transform.IsSetOff())
                {
                    return new XDDFPoint2D(transform.off);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                CT_Point2D xformOff;
                if(value == null)
                {
                    if(transform.IsSetOff())
                    {
                        transform.UnsetOff();
                    }
                    return;
                }
                else if(transform.IsSetOff())
                {
                    xformOff = transform.off;
                }
                else
                {
                    xformOff = transform.AddNewOff();
                }
                xformOff.x = value.X;
                xformOff.y = value.Y;
            }
        }

        public int? Rotation
        {
            get
            {
                if(transform.rotSpecified)
                {
                    return transform.rot;
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
                    transform.rotSpecified = false;
                }
                else
                {
                    transform.rot = value.Value;
                }
            }
        }
    }
}
