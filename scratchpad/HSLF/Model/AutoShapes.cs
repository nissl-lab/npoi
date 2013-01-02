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

using NPOI.ddf.EscherProperties;



/**
 * Stores defInition of auto-shapes.
 * See the Office Drawing 97-2007 Binary Format Specification for details.
 *
 * TODO: follow the spec and define all the auto-shapes
 *
 * @author Yegor Kozlov
 */
public class AutoShapes {
    protected static ShapeOutline[] shapes;


    /**
     * Return shape outline by shape type
     * @param type shape type see {@link ShapeTypes}
     *
     * @return the shape outline
     */
    public static ShapeOutline GetShapeOutline(int type){
        ShapeOutline outline = shapes[type];
        return outline;
    }

    /**
     * Auto-shapes are defined in the [0,21600] coordinate system.
     * We need to transform it into normal slide coordinates
     *
    */
    public static java.awt.Shape transform(java.awt.Shape outline, Rectangle2D anchor){
        AffineTransform at = new AffineTransform();
        at.translate(anchor.GetX(), anchor.GetY());
        at.scale(
                1.0f/21600*anchor.Width,
                1.0f/21600*anchor.Height
        );
        return at.CreateTransformedShape(outline);
    }

    static {
        shapes = new ShapeOutline[255];

        shapes[ShapeTypes.Rectangle] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                Rectangle2D path = new Rectangle2D.Float(0, 0, 21600, 21600);
                return path;
            }
        };

        shapes[ShapeTypes.RoundRectangle] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 5400);
                RoundRectangle2D path = new RoundRectangle2D.Float(0, 0, 21600, 21600, adjval, adjval);
                return path;
            }
        };

        shapes[ShapeTypes.Ellipse] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                Ellipse2D path = new Ellipse2D.Float(0, 0, 21600, 21600);
                return path;
            }
        };

        shapes[ShapeTypes.Diamond] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                GeneralPath path = new GeneralPath();
                path.moveTo(10800, 0);
                path.lineTo(21600, 10800);
                path.lineTo(10800, 21600);
                path.lineTo(0, 10800);
                path.ClosePath();
                return path;
           }
        };

        //m@0,l,21600r21600
        shapes[ShapeTypes.IsocelesTriangle] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 10800);
                GeneralPath path = new GeneralPath();
                path.moveTo(adjval, 0);
                path.lineTo(0, 21600);
                path.lineTo(21600, 21600);
                path.ClosePath();
                return path;
           }
        };

        shapes[ShapeTypes.RightTriangle] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                GeneralPath path = new GeneralPath();
                path.moveTo(0, 0);
                path.lineTo(21600, 21600);
                path.lineTo(0, 21600);
                path.ClosePath();
                return path;
           }
        };

        shapes[ShapeTypes.Parallelogram] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 5400);

                GeneralPath path = new GeneralPath();
                path.moveTo(adjval, 0);
                path.lineTo(21600, 0);
                path.lineTo(21600 - adjval, 21600);
                path.lineTo(0, 21600);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.Trapezoid] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 5400);

                GeneralPath path = new GeneralPath();
                path.moveTo(0, 0);
                path.lineTo(adjval, 21600);
                path.lineTo(21600 - adjval, 21600);
                path.lineTo(21600, 0);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.Hexagon] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 5400);

                GeneralPath path = new GeneralPath();
                path.moveTo(adjval, 0);
                path.lineTo(21600 - adjval, 0);
                path.lineTo(21600, 10800);
                path.lineTo(21600 - adjval, 21600);
                path.lineTo(adjval, 21600);
                path.lineTo(0, 10800);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.Octagon] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 6326);

                GeneralPath path = new GeneralPath();
                path.moveTo(adjval, 0);
                path.lineTo(21600 - adjval, 0);
                path.lineTo(21600, adjval);
                path.lineTo(21600, 21600-adjval);
                path.lineTo(21600-adjval, 21600);
                path.lineTo(adjval, 21600);
                path.lineTo(0, 21600-adjval);
                path.lineTo(0, adjval);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.Plus] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 5400);

                GeneralPath path = new GeneralPath();
                path.moveTo(adjval, 0);
                path.lineTo(21600 - adjval, 0);
                path.lineTo(21600 - adjval, adjval);
                path.lineTo(21600, adjval);
                path.lineTo(21600, 21600-adjval);
                path.lineTo(21600-adjval, 21600-adjval);
                path.lineTo(21600-adjval, 21600);
                path.lineTo(adjval, 21600);
                path.lineTo(adjval, 21600-adjval);
                path.lineTo(0, 21600-adjval);
                path.lineTo(0, adjval);
                path.lineTo(adjval, adjval);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.Pentagon] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){

                GeneralPath path = new GeneralPath();
                path.moveTo(10800, 0);
                path.lineTo(21600, 8259);
                path.lineTo(21600 - 4200, 21600);
                path.lineTo(4200, 21600);
                path.lineTo(0, 8259);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.DownArrow] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                //m0@0 l@1@0 @1,0 @2,0 @2@0,21600@0,10800,21600xe
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 16200);
                int adjval2 = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUST2VALUE, 5400);
                GeneralPath path = new GeneralPath();
                path.moveTo(0, adjval);
                path.lineTo(adjval2, adjval);
                path.lineTo(adjval2, 0);
                path.lineTo(21600-adjval2, 0);
                path.lineTo(21600-adjval2, adjval);
                path.lineTo(21600, adjval);
                path.lineTo(10800, 21600);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.UpArrow] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                //m0@0 l@1@0 @1,21600@2,21600@2@0,21600@0,10800,xe
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 5400);
                int adjval2 = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUST2VALUE, 5400);
                GeneralPath path = new GeneralPath();
                path.moveTo(0, adjval);
                path.lineTo(adjval2, adjval);
                path.lineTo(adjval2, 21600);
                path.lineTo(21600-adjval2, 21600);
                path.lineTo(21600-adjval2, adjval);
                path.lineTo(21600, adjval);
                path.lineTo(10800, 0);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.Arrow] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                //m@0, l@0@1 ,0@1,0@2@0@2@0,21600,21600,10800xe
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 16200);
                int adjval2 = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUST2VALUE, 5400);
                GeneralPath path = new GeneralPath();
                path.moveTo(adjval, 0);
                path.lineTo(adjval, adjval2);
                path.lineTo(0, adjval2);
                path.lineTo(0, 21600-adjval2);
                path.lineTo(adjval, 21600-adjval2);
                path.lineTo(adjval, 21600);
                path.lineTo(21600, 10800);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.LeftArrow] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                //m@0, l@0@1,21600@1,21600@2@0@2@0,21600,,10800xe
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 5400);
                int adjval2 = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUST2VALUE, 5400);
                GeneralPath path = new GeneralPath();
                path.moveTo(adjval, 0);
                path.lineTo(adjval, adjval2);
                path.lineTo(21600, adjval2);
                path.lineTo(21600, 21600-adjval2);
                path.lineTo(adjval, 21600-adjval2);
                path.lineTo(adjval, 21600);
                path.lineTo(0, 10800);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.Can] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                //m10800,qx0@1l0@2qy10800,21600,21600@2l21600@1qy10800,xem0@1qy10800@0,21600@1nfe
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 5400);

                GeneralPath path = new GeneralPath();

                path.Append(new Arc2D.Float(0, 0, 21600, adjval, 0, 180, Arc2D.OPEN), false);
                path.moveTo(0, adjval/2);

                path.lineTo(0, 21600 - adjval/2);
                path.ClosePath();

                path.Append(new Arc2D.Float(0, 21600 - adjval, 21600, adjval, 180, 180, Arc2D.OPEN), false);
                path.moveTo(21600, 21600 - adjval/2);

                path.lineTo(21600, adjval/2);
                path.Append(new Arc2D.Float(0, 0, 21600, adjval, 180, 180, Arc2D.OPEN), false);
                path.moveTo(0, adjval/2);
                path.ClosePath();
                return path;
            }
        };

        shapes[ShapeTypes.LeftBrace] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                //m21600,qx10800@0l10800@2qy0@11,10800@3l10800@1qy21600,21600e
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 1800);
                int adjval2 = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUST2VALUE, 10800);

                GeneralPath path = new GeneralPath();
                path.moveTo(21600, 0);

                path.Append(new Arc2D.Float(10800, 0, 21600, adjval*2, 90, 90, Arc2D.OPEN), false);
                path.moveTo(10800, adjval);

                path.lineTo(10800, adjval2 - adjval);

                path.Append(new Arc2D.Float(-10800, adjval2 - 2*adjval, 21600, adjval*2, 270, 90, Arc2D.OPEN), false);
                path.moveTo(0, adjval2);

                path.Append(new Arc2D.Float(-10800, adjval2, 21600, adjval*2, 0, 90, Arc2D.OPEN), false);
                path.moveTo(10800, adjval2 + adjval);

                path.lineTo(10800, 21600 - adjval);

                path.Append(new Arc2D.Float(10800, 21600 - 2*adjval, 21600, adjval*2, 180, 90, Arc2D.OPEN), false);

                return path;
            }
        };

        shapes[ShapeTypes.RightBrace] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                //m,qx10800@0 l10800@2qy21600@11,10800@3l10800@1qy,21600e
                int adjval = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUSTVALUE, 1800);
                int adjval2 = shape.GetEscherProperty(EscherProperties.GEOMETRY__ADJUST2VALUE, 10800);

                GeneralPath path = new GeneralPath();
                path.moveTo(0, 0);

                path.Append(new Arc2D.Float(-10800, 0, 21600, adjval*2, 0, 90, Arc2D.OPEN), false);
                path.moveTo(10800, adjval);

                path.lineTo(10800, adjval2 - adjval);

                path.Append(new Arc2D.Float(10800, adjval2 - 2*adjval, 21600, adjval*2, 180, 90, Arc2D.OPEN), false);
                path.moveTo(21600, adjval2);

                path.Append(new Arc2D.Float(10800, adjval2, 21600, adjval*2, 90, 90, Arc2D.OPEN), false);
                path.moveTo(10800, adjval2 + adjval);

                path.lineTo(10800, 21600 - adjval);

                path.Append(new Arc2D.Float(-10800, 21600 - 2*adjval, 21600, adjval*2, 270, 90, Arc2D.OPEN), false);

                return path;
            }
        };

        shapes[ShapeTypes.StraightConnector1] = new ShapeOutline(){
            public java.awt.Shape GetOutline(Shape shape){
                return new Line2D.Float(0, 0, 21600, 21600);
            }
        };


    }
}





