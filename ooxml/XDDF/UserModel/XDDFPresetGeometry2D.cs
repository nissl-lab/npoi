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
    using System.Linq;

    public class XDDFPresetGeometry2D
    {
        private CT_PresetGeometry2D geometry;

        public XDDFPresetGeometry2D(CT_PresetGeometry2D geometry)
        {
            this.geometry = geometry;
        }
        public CT_PresetGeometry2D GetXmlObject()
        {
            return geometry;
        }

        public PresetGeometry Geometry
        {
            get
            {
                return PresetGeometryExtensions.ValueOf(geometry.prst);
            }
            set
            {
                geometry.prst = value.ToST_ShapeType();
            }
        }

        public XDDFGeometryGuide AddAdjustValue()
        {
            if(!geometry.IsSetAvLst())
            {
                geometry.AddNewAvLst();
            }
            return new XDDFGeometryGuide(geometry.avLst.AddNewGd());
        }

        public XDDFGeometryGuide InsertAdjustValue(int index)
        {
            if(!geometry.IsSetAvLst())
            {
                geometry.AddNewAvLst();
            }
            return new XDDFGeometryGuide(geometry.avLst.InsertNewGd(index));
        }

        public void RemoveAdjustValue(int index)
        {
            if(geometry.IsSetAvLst())
            {
                geometry.avLst.RemoveGd(index);
            }
        }

        public XDDFGeometryGuide GetAdjustValue(int index)
        {
            if(geometry.IsSetAvLst())
            {
                return new XDDFGeometryGuide(geometry.avLst.GetGdArray(index));
            }
            else
            {
                return null;
            }
        }

        public List<XDDFGeometryGuide> GetAdjustValues()
        {
            if(geometry.IsSetAvLst())
            {
                return geometry.avLst.gd.Select(x => new XDDFGeometryGuide(x)).ToList();
            }
            else
            {
                return new List<XDDFGeometryGuide>();
            }
        }
    }
}
