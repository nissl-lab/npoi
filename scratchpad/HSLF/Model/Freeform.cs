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

namespace NPOI.HSLF.Model;










using NPOI.ddf.EscherArrayProperty;
using NPOI.ddf.EscherContainerRecord;
using NPOI.ddf.EscherOptRecord;
using NPOI.ddf.EscherProperties;
using NPOI.ddf.EscherSimpleProperty;
using NPOI.util.LittleEndian;
using NPOI.util.POILogger;

/**
 * A "Freeform" shape.
 *
 * <p>
 * Shapes Drawn with the "Freeform" tool have cubic bezier curve segments in the smooth sections
 * and straight-line segments in the straight sections. This object closely corresponds to <code>java.awt.geom.GeneralPath</code>.
 * </p>
 * @author Yegor Kozlov
 */
public class Freeform : AutoShape {

    public static byte[] SEGMENTINFO_MOVETO   = new byte[]{0x00, 0x40};
    public static byte[] SEGMENTINFO_LINETO   = new byte[]{0x00, (byte)0xAC};
    public static byte[] SEGMENTINFO_ESCAPE   = new byte[]{0x01, 0x00};
    public static byte[] SEGMENTINFO_ESCAPE2  = new byte[]{0x01, 0x20};
    public static byte[] SEGMENTINFO_CUBICTO  = new byte[]{0x00, (byte)0xAD};
    public static byte[] SEGMENTINFO_CUBICTO2 = new byte[]{0x00, (byte)0xB3}; //OpenOffice inserts 0xB3 instead of 0xAD.
    public static byte[] SEGMENTINFO_CLOSE    = new byte[]{0x01, (byte)0x60};
    public static byte[] SEGMENTINFO_END      = new byte[]{0x00, (byte)0x80};

    /**
     * Create a Freeform object and Initialize it from the supplied Record Container.
     *
     * @param escherRecord       <code>EscherSpContainer</code> Container which holds information about this shape
     * @param parent    the parent of the shape
     */
   protected Freeform(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);

    }

    /**
     * Create a new Freeform. This constructor is used when a new shape is Created.
     *
     * @param parent    the parent of this Shape. For example, if this text box is a cell
     * in a table then the parent is Table.
     */
    public Freeform(Shape parent){
        base(null, parent);
        _escherContainer = CreateSpContainer(ShapeTypes.NotPrimitive, parent is ShapeGroup);
    }

    /**
     * Create a new Freeform. This constructor is used when a new shape is Created.
     *
     */
    public Freeform(){
        this(null);
    }

    /**
     * Set the shape path
     *
     * @param path
     */
    public void SetPath(GeneralPath path)
    {
        Rectangle2D bounds = path.GetBounds2D();
        PathIterator it = path.GetPathIterator(new AffineTransform());

        List<byte[]> segInfo = new List<byte[]>();
        List<Point2D.Double> pntInfo = new List<Point2D.Double>();
        bool IsClosed = false;
        while (!it.IsDone()) {
            double[] vals = new double[6];
            int type = it.currentSegment(vals);
            switch (type) {
                case PathIterator.SEG_MOVETO:
                    pntInfo.Add(new Point2D.Double(vals[0], vals[1]));
                    segInfo.Add(SEGMENTINFO_MOVETO);
                    break;
                case PathIterator.SEG_LINETO:
                    pntInfo.Add(new Point2D.Double(vals[0], vals[1]));
                    segInfo.Add(SEGMENTINFO_LINETO);
                    segInfo.Add(SEGMENTINFO_ESCAPE);
                    break;
                case PathIterator.SEG_CUBICTO:
                    pntInfo.Add(new Point2D.Double(vals[0], vals[1]));
                    pntInfo.Add(new Point2D.Double(vals[2], vals[3]));
                    pntInfo.Add(new Point2D.Double(vals[4], vals[5]));
                    segInfo.Add(SEGMENTINFO_CUBICTO);
                    segInfo.Add(SEGMENTINFO_ESCAPE2);
                    break;
                case PathIterator.SEG_QUADTO:
                    //TODO: figure out how to convert SEG_QUADTO into SEG_CUBICTO
                    logger.log(POILogger.WARN, "SEG_QUADTO is not supported");
                    break;
                case PathIterator.SEG_CLOSE:
                    pntInfo.Add(pntInfo.Get(0));
                    segInfo.Add(SEGMENTINFO_LINETO);
                    segInfo.Add(SEGMENTINFO_ESCAPE);
                    segInfo.Add(SEGMENTINFO_LINETO);
                    segInfo.Add(SEGMENTINFO_CLOSE);
                    isClosed = true;
                    break;
            }

            it.next();
        }
        if(!isClosed) segInfo.Add(SEGMENTINFO_LINETO);
        segInfo.Add(new byte[]{0x00, (byte)0x80});

        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__SHAPEPATH, 0x4));

        EscherArrayProperty verticesProp = new EscherArrayProperty((short)(EscherProperties.GEOMETRY__VERTICES + 0x4000), false, null);
        verticesProp.SetNumberOfElementsInArray(pntInfo.Count);
        verticesProp.SetNumberOfElementsInMemory(pntInfo.Count);
        verticesProp.SetSizeOfElements(0xFFF0);
        for (int i = 0; i < pntInfo.Count; i++) {
            Point2D.Double pnt = pntInfo.Get(i);
            byte[] data = new byte[4];
            LittleEndian.Putshort(data, 0, (short)((pnt.GetX() - bounds.GetX())*MASTER_DPI/POINT_DPI));
            LittleEndian.Putshort(data, 2, (short)((pnt.GetY() - bounds.GetY())*MASTER_DPI/POINT_DPI));
            verticesProp.SetElement(i, data);
        }
        opt.AddEscherProperty(verticesProp);

        EscherArrayProperty segmentsProp = new EscherArrayProperty((short)(EscherProperties.GEOMETRY__SEGMENTINFO + 0x4000), false, null);
        segmentsProp.SetNumberOfElementsInArray(segInfo.Count);
        segmentsProp.SetNumberOfElementsInMemory(segInfo.Count);
        segmentsProp.SetSizeOfElements(0x2);
        for (int i = 0; i < segInfo.Count; i++) {
            byte[] seg = segInfo.Get(i);
            segmentsProp.SetElement(i, seg);
        }
        opt.AddEscherProperty(segmentsProp);

        opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__RIGHT, (int)(bounds.Width*MASTER_DPI/POINT_DPI)));
        opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__BOTTOM, (int)(bounds.Height*MASTER_DPI/POINT_DPI)));

        opt.sortProperties();

        SetAnchor(bounds);
    }

    /**
     * Gets the freeform path
     *
     * @return the freeform path
     */
     public GeneralPath GetPath(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__SHAPEPATH, 0x4));

        EscherArrayProperty verticesProp = (EscherArrayProperty)getEscherProperty(opt, (short)(EscherProperties.GEOMETRY__VERTICES + 0x4000));
        if(verticesProp == null) verticesProp = (EscherArrayProperty)getEscherProperty(opt, EscherProperties.GEOMETRY__VERTICES);

        EscherArrayProperty segmentsProp = (EscherArrayProperty)getEscherProperty(opt, (short)(EscherProperties.GEOMETRY__SEGMENTINFO + 0x4000));
        if(segmentsProp == null) segmentsProp = (EscherArrayProperty)getEscherProperty(opt, EscherProperties.GEOMETRY__SEGMENTINFO);

        //sanity check
        if(verticesProp == null) {
            logger.log(POILogger.WARN, "Freeform is missing GEOMETRY__VERTICES ");
            return null;
        }
        if(segmentsProp == null) {
            logger.log(POILogger.WARN, "Freeform is missing GEOMETRY__SEGMENTINFO ");
            return null;
        }

        GeneralPath path = new GeneralPath();
        int numPoints = verticesProp.GetNumberOfElementsInArray();
        int numSegments = segmentsProp.GetNumberOfElementsInArray();
        for (int i = 0, j = 0; i < numSegments && j < numPoints; i++) {
            byte[] elem = segmentsProp.GetElement(i);
            if(Arrays.Equals(elem, SEGMENTINFO_MOVETO)){
                byte[] p = verticesProp.GetElement(j++);
                short x = LittleEndian.Getshort(p, 0);
                short y = LittleEndian.Getshort(p, 2);
                path.moveTo(
                        ((float)x*POINT_DPI/MASTER_DPI),
                        ((float)y*POINT_DPI/MASTER_DPI));
            } else if (Arrays.Equals(elem, SEGMENTINFO_CUBICTO) || Arrays.Equals(elem, SEGMENTINFO_CUBICTO2)){
                i++;
                byte[] p1 = verticesProp.GetElement(j++);
                short x1 = LittleEndian.Getshort(p1, 0);
                short y1 = LittleEndian.Getshort(p1, 2);
                byte[] p2 = verticesProp.GetElement(j++);
                short x2 = LittleEndian.Getshort(p2, 0);
                short y2 = LittleEndian.Getshort(p2, 2);
                byte[] p3 = verticesProp.GetElement(j++);
                short x3 = LittleEndian.Getshort(p3, 0);
                short y3 = LittleEndian.Getshort(p3, 2);
                path.curveTo(
                        ((float)x1*POINT_DPI/MASTER_DPI), ((float)y1*POINT_DPI/MASTER_DPI),
                        ((float)x2*POINT_DPI/MASTER_DPI), ((float)y2*POINT_DPI/MASTER_DPI),
                        ((float)x3*POINT_DPI/MASTER_DPI), ((float)y3*POINT_DPI/MASTER_DPI));

            } else if (Arrays.Equals(elem, SEGMENTINFO_LINETO)){
                i++;
                byte[] pnext = segmentsProp.GetElement(i);
                if(Arrays.Equals(pnext, SEGMENTINFO_ESCAPE)){
                    if(j + 1 < numPoints){
                        byte[] p = verticesProp.GetElement(j++);
                        short x = LittleEndian.Getshort(p, 0);
                        short y = LittleEndian.Getshort(p, 2);
                        path.lineTo(
                                ((float)x*POINT_DPI/MASTER_DPI), ((float)y*POINT_DPI/MASTER_DPI));
                    }
                } else if (Arrays.Equals(pnext, SEGMENTINFO_CLOSE)){
                    path.ClosePath();
                }
            }
        }
        return path;
    }

    public java.awt.Shape GetOutline(){
        GeneralPath path =  GetPath();
        Rectangle2D anchor = GetAnchor2D();
        Rectangle2D bounds = path.GetBounds2D();
        AffineTransform at = new AffineTransform();
        at.translate(anchor.GetX(), anchor.GetY());
        at.scale(
                anchor.Width/bounds.Width,
                anchor.Height/bounds.Height
        );
        return at.CreateTransformedShape(path);
    }
}





