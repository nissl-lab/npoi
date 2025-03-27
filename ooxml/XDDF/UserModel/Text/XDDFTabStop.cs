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
    using NPOI.OpenXmlFormats.Dml;
    public class XDDFTabStop
    {
        private CT_TextTabStop stop;
        protected XDDFTabStop(CT_TextTabStop stop)
        {
            this.stop = stop;
        }
        protected CT_TextTabStop GetXmlObject()
        {
            return stop;
        }

        public TabAlignment? GetAlignment()
        {
            if(stop.IsSetAlgn())
            {
                return TabAlignmentExtensions.ValueOf(stop.algn);
            }
            else
            {
                return null;
            }
        }

        public void SetAlignment(TabAlignment? align)
        {
            if(align == null)
            {
                if(stop.IsSetAlgn())
                {
                    stop.UnsetAlgn();
                }
            }
            else
            {
                stop.algn = align.Value.ToST_TextTabAlignType();
            }
        }

        public double? GetPosition()
        {
            if(stop.IsSetPos())
            {
                return Units.ToPoints(stop.pos);
            }
            else
            {
                return null;
            }
        }

        public void SetPosition(double? position)
        {
            if(position == null)
            {
                if(stop.IsSetPos())
                {
                    stop.UnsetPos();
                }
            }
            else
            {
                stop.pos = Units.ToEMU(position.Value);
            }
        }
    }
}


