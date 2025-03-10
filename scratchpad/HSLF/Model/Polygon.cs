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

using NPOI.ddf.*;
using NPOI.util.LittleEndian;



/**
 * A simple closed polygon shape
 *
 * @author Yegor Kozlov
 */
public class Polygon : AutoShape {
    /**
     * Create a Polygon object and Initialize it from the supplied Record Container.
     *
     * @param escherRecord       <code>EscherSpContainer</code> Container which holds information about this shape
     * @param parent    the parent of the shape
     */
   protected Polygon(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);

    }

    /**
     * Create a new Polygon. This constructor is used when a new shape is Created.
     *
     * @param parent    the parent of this Shape. For example, if this text box is a cell
     * in a table then the parent is Table.
     */
    public Polygon(Shape parent){
        base(null, parent);
        _escherContainer = CreateSpContainer(ShapeTypes.NotPrimitive, parent is ShapeGroup);
    }

    /**
     * Create a new Polygon. This constructor is used when a new shape is Created.
     *
     */
    public Polygon(){
        this(null);
    }

    /**
     * Set the polygon vertices
     *
     * @param xPoints
     * @param yPoints
     */
    public void SetPoints(float[] xPoints, float[] yPoints)
    {
        float right  = FindBiggest(xPoints);
        float bottom = FindBiggest(yPoints);
        float left   = FindSmallest(xPoints);
        float top    = FindSmallest(yPoints);

        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__RIGHT, (int)((right - left)*POINT_DPI/MASTER_DPI)));
        opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__BOTTOM, (int)((bottom - top)*POINT_DPI/MASTER_DPI)));

        for (int i = 0; i < xPoints.Length; i++) {
            xPoints[i] += -left;
            yPoints[i] += -top;
        }

        int numpoints = xPoints.Length;

        EscherArrayProperty verticesProp = new EscherArrayProperty(EscherProperties.GEOMETRY__VERTICES, false, Array.Empty<byte>() );
        verticesProp.SetNumberOfElementsInArray(numpoints+1);
        verticesProp.SetNumberOfElementsInMemory(numpoints+1);
        verticesProp.SetSizeOfElements(0xFFF0);
        for (int i = 0; i < numpoints; i++)
        {
            byte[] data = new byte[4];
            LittleEndian.Putshort(data, 0, (short)(xPoints[i]*POINT_DPI/MASTER_DPI));
            LittleEndian.Putshort(data, 2, (short)(yPoints[i]*POINT_DPI/MASTER_DPI));
            verticesProp.SetElement(i, data);
        }
        byte[] data = new byte[4];
        LittleEndian.Putshort(data, 0, (short)(xPoints[0]*POINT_DPI/MASTER_DPI));
        LittleEndian.Putshort(data, 2, (short)(yPoints[0]*POINT_DPI/MASTER_DPI));
        verticesProp.SetElement(numpoints, data);
        opt.AddEscherProperty(verticesProp);

        EscherArrayProperty segmentsProp = new EscherArrayProperty(EscherProperties.GEOMETRY__SEGMENTINFO, false, null );
        segmentsProp.SetSizeOfElements(0x0002);
        segmentsProp.SetNumberOfElementsInArray(numpoints * 2 + 4);
        segmentsProp.SetNumberOfElementsInMemory(numpoints * 2 + 4);
        segmentsProp.SetElement(0, new byte[] { (byte)0x00, (byte)0x40 } );
        segmentsProp.SetElement(1, new byte[] { (byte)0x00, (byte)0xAC } );
        for (int i = 0; i < numpoints; i++)
        {
            segmentsProp.SetElement(2 + i * 2, new byte[] { (byte)0x01, (byte)0x00 } );
            segmentsProp.SetElement(3 + i * 2, new byte[] { (byte)0x00, (byte)0xAC } );
        }
        segmentsProp.SetElement(segmentsProp.GetNumberOfElementsInArray() - 2, new byte[] { (byte)0x01, (byte)0x60 } );
        segmentsProp.SetElement(segmentsProp.GetNumberOfElementsInArray() - 1, new byte[] { (byte)0x00, (byte)0x80 } );
        opt.AddEscherProperty(segmentsProp);

        opt.sortProperties();
    }

    /**
     * Set the polygon vertices
     *
     * @param points the polygon vertices
     */
      public void SetPoints(Point2D[] points)
     {
        float[] xpoints = new float[points.Length];
        float[] ypoints = new float[points.Length];
        for (int i = 0; i < points.Length; i++) {
            xpoints[i] = (float)points[i].GetX();
            ypoints[i] = (float)points[i].GetY();

        }

        SetPoints(xpoints, ypoints);
    }

    private float FindBiggest( float[] values )
    {
        float result = Float.MinValue;
        for ( int i = 0; i < values.Length; i++ )
        {
            if (values[i] > result)
                result = values[i];
        }
        return result;
    }

    private float FindSmallest( float[] values )
    {
        float result = Float.MaxValue;
        for ( int i = 0; i < values.Length; i++ )
        {
            if (values[i] < result)
                result = values[i];
        }
        return result;
    }


}





