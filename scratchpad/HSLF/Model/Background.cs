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

using NPOI.ddf.EscherContainerRecord;
using NPOI.HSLF.usermodel.PictureData;
using NPOI.HSLF.blip.Bitmap;
using NPOI.util.POILogger;

using javax.imageio.ImageIO;




/**
 * Background shape
 *
 * @author Yegor Kozlov
 */
public class Background : Shape {

    protected Background(EscherContainerRecord escherRecord, Shape parent) {
        base(escherRecord, parent);
    }

    protected EscherContainerRecord CreateSpContainer(bool IsChild) {
        return null;
    }

    public void Draw(Graphics2D graphics) {
        Fill f = GetFill();
        Dimension pg = Sheet.GetSlideShow().GetPageSize();
        Rectangle anchor = new Rectangle(0, 0, pg.width, pg.height);
        switch (f.GetFillType()) {
            case Fill.FILL_SOLID:
                Color color = f.GetForegroundColor();
                graphics.SetPaint(color);
                graphics.Fill(anchor);
                break;
            case Fill.FILL_PICTURE:
                PictureData data = f.GetPictureData();
                if (data is Bitmap) {
                    BufferedImage img = null;
                    try {
                        img = ImageIO.Read(new MemoryStream(data.Data));
                    } catch (Exception e) {
                        logger.log(POILogger.WARN, "ImageIO failed to create image. image.type: " + data.GetType());
                        return;
                    }
                    Image scaledImg = img.GetScaledInstance(anchor.width, anchor.height, Image.SCALE_SMOOTH);
                    graphics.DrawImage(scaledImg, anchor.x, anchor.y, null);

                }
                break;
            default:
                logger.log(POILogger.WARN, "unsuported fill type: " + f.GetFillType());
                break;
        }
    }
}





