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
    public class XDDFNormalAutoFit : IXDDFAutoFit
    {
        private CT_TextNormalAutofit autofit;

        public XDDFNormalAutoFit()
            : this(new CT_TextNormalAutofit())
        {
            
        }
        protected XDDFNormalAutoFit(CT_TextNormalAutofit autofit)
        {
            this.autofit = autofit;
        }
        protected CT_TextNormalAutofit GetXmlobject()
        {
            return autofit;
        }
        public int GetFontScale()
        {
            if(autofit.IsSetFontScale())
            {
                return autofit.fontScale;
            }
            else
            {
                return 100_000;
            }
        }

        public void SetFontScale(int? value)
        {
            if(value == null)
            {
                if(autofit.IsSetFontScale())
                {
                    autofit.UnsetFontScale();
                }
            }
            else
            {
                autofit.fontScale = value.Value;
            }

        }

        public int GetLineSpaceReduction()
        {
            if(autofit.IsSetLnSpcReduction())
            {
                return autofit.lnSpcReduction;
            }
            else
            {
                return 0;
            }
        }

        public void setLineSpaceReduction(int? value)
        {
            if(value == null)
            {
                if(autofit.IsSetLnSpcReduction())
                {
                    autofit.UnsetLnSpcReduction();
                }
            }
            else
            {
                autofit.lnSpcReduction = value.Value;
            }
        }
    }
}