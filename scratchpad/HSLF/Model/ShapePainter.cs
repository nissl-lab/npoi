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


using NPOI.util.POILogger;
using NPOI.util.POILogFactory;




/**
 * Paint a shape into java.awt.Graphics2D
 *
 * @author Yegor Kozlov
 */
public class ShapePainter {
    protected static POILogger logger = POILogFactory.GetLogger(ShapePainter.class);

    public static void paint(SimpleShape shape, Graphics2D graphics){
        Rectangle2D anchor = shape.GetLogicalAnchor2D();
        java.awt.Shape outline = shape.GetOutline();

        //flip vertical
        if(shape.GetFlipVertical()){
            graphics.translate(anchor.GetX(), anchor.GetY() + anchor.Height);
            graphics.scale(1, -1);
            graphics.translate(-anchor.GetX(), -anchor.GetY());
        }
        //flip horizontal
        if(shape.GetFlipHorizontal()){
            graphics.translate(anchor.GetX() + anchor.Width, anchor.GetY());
            graphics.scale(-1, 1);
            graphics.translate(-anchor.GetX() , -anchor.GetY());
        }

        //rotate transform
        double angle = shape.GetRotation();

        if(angle != 0){
            double centerX = anchor.GetX() + anchor.Width/2;
            double centerY = anchor.GetY() + anchor.Height/2;

            graphics.translate(centerX, centerY);
            graphics.rotate(Math.ToRadians(angle));
            graphics.translate(-centerX, -centerY);
        }

        //fill
        Color FillColor = shape.GetFill().GetForegroundColor();
        if (FillColor != null) {
            //TODO: implement gradient and texture fill patterns
            graphics.SetPaint(FillColor);
            graphics.Fill(outline);
        }

        //border
        Color lineColor = shape.GetLineColor();
        if (lineColor != null){
            graphics.SetPaint(lineColor);
            float width = (float)shape.GetLineWidth();

            int dashing = shape.GetLineDashing();
            //TODO: implement more dashing styles
            float[] dashptrn = null;
            switch(dashing){
                case Line.PEN_SOLID:
                    dashptrn = null;
                    break;
                case Line.PEN_PS_DASH:
                    dashptrn = new float[]{width, width};
                    break;
                case Line.PEN_DOTGEL:
                    dashptrn = new float[]{width*4, width*3};
                    break;
               default:
                    logger.log(POILogger.WARN, "unsupported dashing: " + dashing);
                    dashptrn = new float[]{width, width};
                    break;
            }

            Stroke stroke = new BasicStroke(width, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER, 10.0f, dashptrn, 0.0f);
            graphics.SetStroke(stroke);
            graphics.Draw(outline);
        }
    }
}





