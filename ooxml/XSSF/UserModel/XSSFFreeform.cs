/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace NPOI.XSSF.UserModel
{
    public class XSSFFreeform : XSSFShape
    {
        /**
         * A default instance of CTShape used for creating new shapes.
         */
        private static CT_Shape prototype = null;

        /**
           *  Xml bean that stores properties of this shape
           */
        private CT_Shape ctShape;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Drawing"></param>
        /// <param name="ctShape"></param>
        protected internal XSSFFreeform(XSSFDrawing Drawing, CT_Shape ctShape)
        {
            this.drawing = Drawing;
            this.ctShape = ctShape;
        }
        /**
         * Prototype with the default structure of a new auto-shape.
         */
        protected internal static CT_Shape Prototype()
        {
            CT_Shape shape = new CT_Shape();

            CT_ShapeNonVisual nv = shape.AddNewNvSpPr();
            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_NonVisualDrawingProps nvp = nv.AddNewCNvPr();
            nvp.id = (/*setter*/1);
            nvp.name = (/*setter*/"Freeform 1");
            nv.AddNewCNvSpPr();

            NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties sp = shape.AddNewSpPr();
            CT_Transform2D t2d = sp.AddNewXfrm();
            CT_PositiveSize2D p1 = t2d.AddNewExt();
            p1.cx = (/*setter*/0);
            p1.cy = (/*setter*/0);
            CT_Point2D p2 = t2d.AddNewOff();
            p2.x = (/*setter*/0);
            p2.y = (/*setter*/0);

            CT_CustomGeometry2D geom = sp.AddNewCustGeom();

            prototype = shape;

            return prototype;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public CT_Shape GetCTShape()
        {
            return ctShape;
        }

        /// <summary>
        /// 
        /// </summary>
        public override uint ID
        {
            get
            {
                return ctShape.nvSpPr.cNvPr.id;
            }
        }

        /**
         * Returns the simple shape name.
         * @return name of the simple shape
         */
        public override String Name
        {
            get
            {
                return ctShape.nvSpPr.cNvPr.name;
            }
            set
            {
                ctShape.nvSpPr.cNvPr.name = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        protected internal override NPOI.OpenXmlFormats.Dml.Spreadsheet.CT_ShapeProperties GetShapeProperties()
        {
            return ctShape.spPr;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="BFF"></param>
        public void Build(BuildFreeForm BFF)
        {
            //-----
            ctShape.spPr.xfrm.off.x = (long)BFF.Min.x;
            ctShape.spPr.xfrm.off.y = (long)BFF.Min.y;
            ctShape.spPr.xfrm.ext.cx = (long)BFF.Width;
            ctShape.spPr.xfrm.ext.cy = (long)BFF.Height;

            //-----
            ctShape.spPr.custGeom.avLst = new CT_GeomGuideList();
            //-----
            ctShape.spPr.custGeom.ahLst = new List<object>();
            //-----
            ctShape.spPr.custGeom.rect = new CT_GeomRect();
            ctShape.spPr.custGeom.rect.l = "l";
            ctShape.spPr.custGeom.rect.t = "t";
            ctShape.spPr.custGeom.rect.b = "b";
            ctShape.spPr.custGeom.rect.r = "r";

            //-----
            ctShape.spPr.custGeom.gdLst = new CT_GeomGuideList();
            ctShape.spPr.custGeom.gdLst.gd = new List<CT_GeomGuide>();
            ctShape.spPr.custGeom.cxnLst = new CT_ConnectionSiteList();
            ctShape.spPr.custGeom.cxnLst.cxn = new List<CT_ConnectionSite>();
            int ct = 0;
            foreach(var c in BFF.Points)
            {
                CT_GeomGuide gg;
                gg = new CT_GeomGuide {
                    name = $"connsiteX{ct}",
                    fmla = $"*/ {c.x} w {BFF.Width}"
                };
                ctShape.spPr.custGeom.gdLst.gd.Add(gg);
                gg = new CT_GeomGuide {
                    name = $"connsiteY{ct}",
                    fmla = $"*/ {c.y} h {BFF.Height}"
                };
                ctShape.spPr.custGeom.gdLst.gd.Add(gg);

                //-----
                CT_ConnectionSite cs;
                cs = new CT_ConnectionSite();
                cs.pos = new CT_AdjPoint2D();
                cs.ang = "0";
                cs.pos.x = $"connsiteX{ct}";
                cs.pos.y = $"connsiteY{ct}";
                ctShape.spPr.custGeom.cxnLst.cxn.Add(cs);
                ct++;
            }

            //-----
            ctShape.spPr.custGeom.pathLst = new CT_Path2DList();
            ctShape.spPr.custGeom.pathLst.path = new List<CT_Path2D>();
            var ctp = new CT_Path2D();
            ctp.moveto = new CT_Path2DMoveTo();
            ctp.moveto.pt = new CT_AdjPoint2D();
            ctp.w = (long)BFF.Width;
            ctp.h = (long)BFF.Height;

            //-----
            ctp.moveto.pt.x = $"{BFF.Points.First.Value.x}";
            ctp.moveto.pt.y = $"{BFF.Points.First.Value.y}";
            for(var bze = BFF.ControlPoints.First; bze!=null; )
            {
                CT_Path2DCubicBezierTo cubBze = new CT_Path2DCubicBezierTo();
                cubBze.pt = new List<CT_AdjPoint2D>();
                for(int i = 0; i < 3; i++)
                {
                    CT_AdjPoint2D pt;
                    pt = new CT_AdjPoint2D();
                    pt.x = $"{bze.Value.x}";
                    pt.y = $"{bze.Value.y}";
                    cubBze.pt.Add(pt);
                    bze = bze.Next;
                }
                ctp.cubicBezTo.Add(cubBze);
            }
            ctShape.spPr.custGeom.pathLst.path.Add(ctp);
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class BuildFreeForm
    {
        public          DblVect2D  mMin;
        public          DblVect2D  mMax;

        protected LinkedList<DblVect2D> mPoints { get; }
        protected LinkedList<DblVect2D> mInterVect{ get; }         //intermediate vector
        protected LinkedList<DblVect2D> mVect { get; set; }

        public DblVect2D Min { get => mMin; }
        public DblVect2D Max { get => mMax; }
        public double Width { get => Math.Abs(mMax.x - mMin.x); }
        public double Height { get => Math.Abs(mMax.y - mMin.y); }
        public LinkedList<Coords> Points
        {
            get
            {
                LinkedList<Coords> res = new LinkedList<Coords>();
                for(var nd = mPoints.First; nd!=null; nd = nd.Next)
                {
                    res.AddLast(new Coords((long)(nd.Value.x-mMin.x), (long)(nd.Value.y-mMin.y)));
                }

                return res;
            }
        }

        public LinkedList<Coords> ControlPoints
        {
            get
            {
                LinkedList<Coords> res = new LinkedList<Coords>();
                double x = Points.First.Value.x;
                double y = Points.First.Value.y;
                LinkedListNode<DblVect2D> nd = mVect.First;

                for(; nd!=null; nd=nd.Next)
                {
                    res.AddLast(new Coords((long)(x+nd.Value.x), (long)(y+nd.Value.y)));
                    nd = nd.Next.Next;
                    x += nd.Value.x;
                    y += nd.Value.y;
                    res.AddLast(new Coords((long)(x+nd.Previous.Value.x)
                                         , (long)(y+nd.Previous.Value.y)));
                    res.AddLast(new Coords((long)x, (long)y));
                }

                return res;
            }
        }

        public BuildFreeForm()
        {
            mMin = new DblVect2D(double.MaxValue, double.MaxValue);
            mMax = new DblVect2D(double.MinValue, double.MinValue);
            mPoints = new LinkedList<DblVect2D>();
            mInterVect = new LinkedList<DblVect2D>();
        }

        public void AddNode(Coords Pt)
        {
            mMin.Min(DblVect2D.Conv(Pt));
            mMax.Max(DblVect2D.Conv(Pt));

            var nd = mPoints.AddLast(DblVect2D.Conv(Pt));
            if(nd.Previous != null)
            {
                mInterVect.AddLast(nd.Value - nd.Previous.Value);
            }
        }

        public bool Build()
        {
            if(mInterVect.Count<2)
            {
                return false;
            }

            DblVect2D cv;                              //control vector
            LinkedListNode<DblVect2D> nd;
            mVect = new LinkedList<DblVect2D>();       //new control vectors

            mVect.AddFirst(new DblVect2D(mInterVect.First.Value.x
                                       , mInterVect.First.Value.y));
            for(nd = mInterVect.First.Next; nd != null; nd = nd.Next)
            {
                cv = nd.Previous.Value + nd.Value;
                double ip = 1;   //Relevance(nd.Previous.Value, nd.Value);
                mVect.AddBefore(mVect.Last, cv * (Relevance(cv, nd.Previous.Value) * -ip));
                mVect.AddLast(cv * (Relevance(cv, nd.Value) * ip));
                mVect.AddLast(new DblVect2D(nd.Value.x, nd.Value.y));
            }

            nd = mVect.First;
            cv = nd.Value + nd.Next.Value;
            mVect.AddBefore(nd, cv * Relevance(cv, nd.Value));

            nd = mVect.Last;
            cv = nd.Value - nd.Previous.Value;
            mVect.AddBefore(nd, cv * (Relevance(cv, nd.Value) * -1));

            return true;
        }

        protected double Relevance(DblVect2D Base, DblVect2D Val)
        {
            return Base.InnerProduct(Val) * -0.35 + 0.4;
        }
    }
}
