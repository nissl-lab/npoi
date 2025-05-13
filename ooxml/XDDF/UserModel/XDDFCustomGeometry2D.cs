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
    using System.Linq;

    public class XDDFCustomGeometry2D
    {
        private CT_CustomGeometry2D geometry;

        public XDDFCustomGeometry2D(CT_CustomGeometry2D geometry)
        {
            this.geometry = geometry;
        }
        public CT_CustomGeometry2D GetXmlObject()
        {
            return geometry;
        }

        public XDDFGeometryRectangle Rectangle
        {
            get
            {
                if(geometry.IsSetRect())
                {
                    return new XDDFGeometryRectangle(geometry.rect);
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
                    if(geometry.IsSetRect())
                    {
                        geometry.UnsetRect();
                    }
                }
                else
                {
                    geometry.rect = value.GetXmlObject();
                }
            }
        }

        public XDDFAdjustHandlePolar AddPolarAdjustHandle()
        {
            if(!geometry.IsSetAhLst())
            {
                geometry.AddNewAhLst();
            }
            return new XDDFAdjustHandlePolar(geometry.ahLst.AddNewAhPolar());
        }

        public XDDFAdjustHandlePolar InsertPolarAdjustHandle(int index)
        {
            if(!geometry.IsSetAhLst())
            {
                geometry.AddNewAhLst();
            }
            return new XDDFAdjustHandlePolar(geometry.ahLst.InsertNewAhPolar(index));
        }

        public void removePolarAdjustHandle(int index)
        {
            if(geometry.IsSetAhLst())
            {
                geometry.ahLst.RemoveAhPolar(index);
            }
        }

        public XDDFAdjustHandlePolar GetPolarAdjustHandle(int index)
        {
            if(geometry.IsSetAhLst())
            {
                return new XDDFAdjustHandlePolar(geometry.ahLst.GetAhPolarArray(index));
            }
            else
            {
                return null;
            }
        }

        public List<XDDFAdjustHandlePolar> GetPolarAdjustHandles()
        {
            if(geometry.IsSetAhLst())
            {
                return [.. geometry.ahLst.GetAhPolarList().Select(x=>new XDDFAdjustHandlePolar(x))];
            }
            else
            {
                return new List<XDDFAdjustHandlePolar>();
            }
        }

        public XDDFAdjustHandleXY AddXYAdjustHandle()
        {
            if(!geometry.IsSetAhLst())
            {
                geometry.AddNewAhLst();
            }
            return new XDDFAdjustHandleXY(geometry.ahLst.AddNewAhXY());
        }

        public XDDFAdjustHandleXY insertXYAdjustHandle(int index)
        {
            if(!geometry.IsSetAhLst())
            {
                geometry.AddNewAhLst();
            }
            return new XDDFAdjustHandleXY(geometry.ahLst.InsertNewAhXY(index));
        }

        public void removeXYAdjustHandle(int index)
        {
            if(geometry.IsSetAhLst())
            {
                geometry.ahLst.RemoveAhXY(index);
            }
        }

        public XDDFAdjustHandleXY GetXYAdjustHandle(int index)
        {
            if(geometry.IsSetAhLst())
            {
                return new XDDFAdjustHandleXY(geometry.ahLst.GetAhXYArray(index));
            }
            else
            {
                return null;
            }
        }

        public List<XDDFAdjustHandleXY> GetXYAdjustHandles()
        {
            if(geometry.IsSetAhLst())
            {
                return [.. geometry.ahLst.GetAhXYList().Select(x=>new XDDFAdjustHandleXY(x))];
            }
            else
            {
                return new List<XDDFAdjustHandleXY>();
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

        public XDDFGeometryGuide insertAdjustValue(int index)
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
                return [.. geometry.avLst.gd.Select(x=> new XDDFGeometryGuide(x))];
            }
            else
            {
                return new List<XDDFGeometryGuide>();
            }
        }

        public XDDFConnectionSite AddConnectionSite()
        {
            if(!geometry.IsSetCxnLst())
            {
                geometry.AddNewCxnLst();
            }
            return new XDDFConnectionSite(geometry.cxnLst.AddNewCxn());
        }

        public XDDFConnectionSite InsertConnectionSite(int index)
        {
            if(!geometry.IsSetCxnLst())
            {
                geometry.AddNewCxnLst();
            }
            return new XDDFConnectionSite(geometry.cxnLst.InsertNewCxn(index));
        }

        public void RemoveConnectionSite(int index)
        {
            if(geometry.IsSetCxnLst())
            {
                geometry.cxnLst.RemoveCxn(index);
            }
        }

        public XDDFConnectionSite GetConnectionSite(int index)
        {
            if(geometry.IsSetCxnLst())
            {
                return new XDDFConnectionSite(geometry.cxnLst.GetCxnArray(index));
            }
            else
            {
                return null;
            }
        }

        public List<XDDFConnectionSite> GetConnectionSites()
        {
            if(geometry.IsSetCxnLst())
            {
                return [.. geometry.cxnLst.cxn.Select(x=>new XDDFConnectionSite(x))];
            }
            else
            {
                return new List<XDDFConnectionSite>();
            }
        }

        public XDDFGeometryGuide AddGuide()
        {
            if(!geometry.IsSetGdLst())
            {
                geometry.AddNewGdLst();
            }
            return new XDDFGeometryGuide(geometry.gdLst.AddNewGd());
        }

        public XDDFGeometryGuide InsertGuide(int index)
        {
            if(!geometry.IsSetGdLst())
            {
                geometry.AddNewGdLst();
            }
            return new XDDFGeometryGuide(geometry.gdLst.InsertNewGd(index));
        }

        public void removeGuide(int index)
        {
            if(geometry.IsSetGdLst())
            {
                geometry.gdLst.RemoveGd(index);
            }
        }

        public XDDFGeometryGuide GetGuide(int index)
        {
            if(geometry.IsSetGdLst())
            {
                return new XDDFGeometryGuide(geometry.gdLst.GetGdArray(index));
            }
            else
            {
                return null;
            }
        }

        public List<XDDFGeometryGuide> GetGuides()
        {
            if(geometry.IsSetGdLst())
            {
                return [.. geometry.gdLst.gd.Select(x=>new XDDFGeometryGuide(x))];
            }
            else
            {
                return new List<XDDFGeometryGuide>();
            }
        }

        public XDDFPath AddNewPath()
        {
            return new XDDFPath(geometry.pathLst.AddNewPath());
        }

        public XDDFPath InsertNewPath(int index)
        {
            return new XDDFPath(geometry.pathLst.InsertNewPath(index));
        }

        public void RemovePath(int index)
        {
            geometry.pathLst.RemovePath(index);
        }

        public XDDFPath GetPath(int index)
        {
            return new XDDFPath(geometry.pathLst.GetPathArray(index));
        }

        public List<XDDFPath> GetPaths()
        {
            return [.. geometry.pathLst.path.Select(x=>new XDDFPath(x))];
        }
    }
}
