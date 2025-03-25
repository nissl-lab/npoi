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
    public class XDDFTileInfoProperties
    {
        private CT_TileInfoProperties props;

        public XDDFTileInfoProperties(CT_TileInfoProperties properties)
        {
            this.props = properties;
        }
        public CT_TileInfoProperties GetXmlobject()
        {
            return props;
        }

        public RectangleAlignment? Alignment
        {
            get 
            {
                if(props.algnSpecified)
                {
                    return RectangleAlignmentExtensions.ValueOf(props.algn);
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
                    props.algn = value.Value.ToST_RectAlignment();
                }
            }
        }

        public TileFlipMode? FlipMode
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

        public int? Sx
        {
            get => props.sxSpecified ? props.sx : null;
            set
            {
                if(value == null)
                {
                    props.sxSpecified = false;
                }
                else
                {
                    props.sx = value.Value;
                }
            }
        }

        public int? Sy
        {
            get => props.sySpecified ? props.sy : null;
            set
            {
                if(value == null)
                {
                    props.sySpecified = false;
                }
                else
                {
                    props.sy = value.Value;
                }
            }
        }

        public long? Tx
        {
            get => props.txSpecified ? props.tx : null;
            set
            {
                if(value == null)
                {
                    props.txSpecified = false;
                }
                else
                {
                    props.tx = value.Value;
                }
            }
        }
        
        public long? Ty
        {
            get
            {
                if(props.tySpecified)
                {
                    return props.ty;
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
                    props.tySpecified = false;
                }
                else
                {
                    props.ty = value.Value;
                }
            }
        }
    }
}
