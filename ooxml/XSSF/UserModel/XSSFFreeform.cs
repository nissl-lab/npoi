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
            ctShape.spPr.xfrm.off.x = BFF.Left;
            ctShape.spPr.xfrm.off.y = BFF.Top;
            ctShape.spPr.xfrm.ext.cx = BFF.Width;
            ctShape.spPr.xfrm.ext.cy = BFF.Height;

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
            ctp.w = BFF.Width;
            ctp.h = BFF.Height;

            //-----
            ctp.moveto.pt.x = $"{Math.Floor(BFF.mCtrlPoint.First.Value.x - BFF.dLeft)}";
            ctp.moveto.pt.y = $"{Math.Floor(BFF.mCtrlPoint.First.Value.y - BFF.dTop)}";
            for(var bze = BFF.mCtrlPoint.First.Next; bze!=null;)
            {
                CT_Path2DCubicBezierTo cubBze = new CT_Path2DCubicBezierTo();
                cubBze.pt = new List<CT_AdjPoint2D>();
                for(int i = 0; i < 3; i++)
                {
                    CT_AdjPoint2D pt;
                    pt = new CT_AdjPoint2D();
                    pt.x = $"{Math.Floor(bze.Value.x - BFF.dLeft)}";
                    pt.y = $"{Math.Floor(bze.Value.y - BFF.dTop)}";
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
        private DblVect2D  mMin;
        private DblVect2D  mMax;

        public long Left
        {
            get { return (long )Math.Round(mMin.x); }
        }
        public long Rigth
        {
            get { return (long) Math.Round(mMax.x); }
        }
        public long Top
        {
            get { return (long)Math.Round(mMin.y); }
        }
        public long Bottom
        {
            get { return (long)Math.Round(mMax.y); }
        }
        internal double dLeft
        {
            get => mMin.x;
        }
        internal double dTop
        {
            get => mMin.y;
        }

        protected LinkedList<DblVect2D> mPoints { get; }                //coordinate
        protected LinkedList<DblVect2D> mInterVect { get; }             //intermediate vector

        public LinkedList<DblVect2D> mCtrlPoint { get; private set; }   //Control point coordinate

        public long Width   { get => Math.Abs(Rigth - Left); }
        public long Height  { get => Math.Abs(Top - Bottom); }
        public LinkedList<Coords> Points
        {
            get
            {
                LinkedList<Coords> res = new LinkedList<Coords>();
                for(var nd = mPoints.First; nd!=null; nd = nd.Next)
                {
                    res.AddLast(new Coords((long) (nd.Value.x-Left)
                                         , (long) (nd.Value.y-Top)));
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

            //========== Calculate vector ==========
            LinkedList<DblVect2D> cv = new LinkedList<DblVect2D>();     //control vectors
            LinkedListNode<DblVect2D> nd;
            DblVect2D hv;                                               //handle vector

            cv.AddFirst(new DblVect2D(mInterVect.First.Value.x
                                    , mInterVect.First.Value.y));
            for(nd = mInterVect.First.Next; nd != null; nd = nd.Next)
            {
                hv = nd.Previous.Value + nd.Value;
                cv.AddBefore(cv.Last, hv / -6);
                cv.AddLast(hv / 6);
                cv.AddLast(new DblVect2D(nd.Value.x, nd.Value.y));
            }

            nd = cv.First;
            hv = nd.Value + nd.Next.Value;
            cv.AddBefore(nd, hv / 2);

            nd = cv.Last;
            hv = nd.Value - nd.Previous.Value;
            cv.AddBefore(nd, hv / -2);

            //========== Convert vector to coordinate ==========
            mCtrlPoint = new LinkedList<DblVect2D>();
            double x = mPoints.First.Value.x;
            double y = mPoints.First.Value.y;
            mCtrlPoint.AddLast(new DblVect2D(x, y));
            nd = cv.First;

            for(; nd!=null; nd=nd.Next)
            {
                mCtrlPoint.AddLast(new DblVect2D(x+nd.Value.x, y+nd.Value.y));
                nd = nd.Next.Next;
                x += nd.Value.x;
                y += nd.Value.y;
                mCtrlPoint.AddLast(new DblVect2D(x+nd.Previous.Value.x
                                               , y+nd.Previous.Value.y));
                mCtrlPoint.AddLast(new DblVect2D((long) x, (long) y));
            }
            //int i = 0;
            //for(var v = mCtrlPoint.First; v != null; v=v.Next)
            //{
            //    Console.WriteLine($"{i++:00}:{v.Value.x},{v.Value.y}");
            //}
            //========== Build Rectangle ==========
            mMin = new DblVect2D(double.MaxValue, double.MaxValue);
            mMax = new DblVect2D(double.MinValue, double.MinValue);
            for(var v = mCtrlPoint.First; v.Next != null;)
            {
                List<DblVect2D> point = new();
                point.Add(v.Value);
                v = v.Next;
                point.Add(v.Value);
                v = v.Next;
                point.Add(v.Value);
                v = v.Next;
                point.Add(v.Value);
                for(double t = 0.0; t < 1.0; t += 0.0025)
                {
                    DblVect2D bez = Bezier(point, t);
                    mMin.Min(bez);
                    mMax.Max(bez);
                    //Console.WriteLine($"{t:#0.00}:{bez.x:0000000.00},{bez.y:0000000.00}");
                }
            }

            return true;
        }

        /// <summary>
        /// P(t)=sigma(k=0..N,nCk*t^k*(1-t)^(N-k)Pk)
        /// </summary>
        /// <param name="p"></param>
        private DblVect2D Bezier(List<DblVect2D> p, double t)
        {
            int n = p.Count - 1;
            DblVect2D res = new(0.0,0.0);

            for(int k = 0; k<p.Count; k++)
            {
                double a = BinomialCoefficients(n,k) * Math.Pow(t, k) * Math.Pow(1-t, n-k);
                DblVect2D val = new( a * p[k].x, a * p[k].y );
                res.Add(val);
                //Console.WriteLine($"{t},{k},{BinomialCoefficients(n, k)},{Math.Pow(t, k)},{Math.Pow(1-t, n-k)},{p[k].x},{p[k].y}");
            }
            return res;
        }

        /// <summary>
        /// nCk
        /// </summary>
        /// <param name="n"></param>
        /// <param name="k"></param>
        /// <returns></returns>
        private long BinomialCoefficients(int n, int k)
        {
            return Factorial(n)/(Factorial(k)*Factorial(n-k));
        }

        /// <summary>
        /// n! (n>=0)
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        private long Factorial(int n)
        {
            return n < 1 ? 1L : n * Factorial(n-1);
        }
    }
}
