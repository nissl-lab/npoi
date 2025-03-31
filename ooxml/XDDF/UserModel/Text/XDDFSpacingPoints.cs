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
    public class XDDFSpacingPoints : XDDFSpacing
    {
        private CT_TextSpacingPoint points;

        public XDDFSpacingPoints(double value)
                : this(new CT_TextSpacing(), new CT_TextSpacingPoint())
        {

            if(spacing.IsSetSpcPct())
            {
                spacing.UnsetSpcPct();
            }
            spacing.spcPts = points;
            SetPoints(value);
        }
        protected XDDFSpacingPoints(CT_TextSpacing parent, CT_TextSpacingPoint points)
                : base(parent)
        {
            ;
            this.points = points;
        }
        public override Kind Type
        {
            get
            {
                return Kind.Points;
            }
        }

        public double GetPoints()
        {
            return points.val * 0.01;
        }

        public void SetPoints(double value)
        {
            points.val = (int) (100 * value);
        }
    }
}


