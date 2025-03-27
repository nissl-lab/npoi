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
    public class XDDFSpacingPercent : XDDFSpacing
    {
        private CT_TextSpacingPercent percent;
        private double scale;

        public XDDFSpacingPercent(double value)
            : this(new CT_TextSpacing(), new CT_TextSpacingPercent(), null)
        {
            
            if(spacing.IsSetSpcPts())
            {
                spacing.UnsetSpcPts();
            }
            spacing.spcPct = percent;
            SetPercent(value);
        }
        protected XDDFSpacingPercent(CT_TextSpacing parent, CT_TextSpacingPercent percent, double? scale)
            : base(parent)
        {
            
            this.percent = percent;
            this.scale = (scale == null) ? 0.001 : scale.Value * 0.001;
        }
        public override Kind Type
        {
            get
            {
                return Kind.Percent;
            }
            
        }

        public double GetPercent()
        {
            return percent.val * scale;
        }

        public void SetPercent(double value)
        {
            percent.val = (int) (1000 * value);
        }
    }
}


