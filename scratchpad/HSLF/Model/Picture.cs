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
using NPOI.HSLF.usermodel.PictureData;
using NPOI.HSLF.usermodel.SlideShow;
using NPOI.HSLF.record.Document;
using NPOI.HSLF.blip.Bitmap;
using NPOI.HSLF.exceptions.HSLFException;
using NPOI.util.POILogger;

using javax.imageio.ImageIO;









/**
 * Represents a picture in a PowerPoint document.
 *
 * @author Yegor Kozlov
 */
public class Picture : SimpleShape {

    /**
    *  Windows Enhanced Metafile (EMF)
    */
    public static int EMF = 2;

    /**
    *  Windows Metafile (WMF)
    */
    public static int WMF = 3;

    /**
    * Macintosh PICT
    */
    public static int PICT = 4;

    /**
    *  JPEG
    */
    public static int JPEG = 5;

    /**
    *  PNG
    */
    public static int PNG = 6;

    /**
     * Windows DIB (BMP)
     */
    public static byte DIB = 7;

    /**
     * Create a new <code>Picture</code>
     *
    * @param idx the index of the picture
     */
    public Picture(int idx){
        this(idx, null);
    }

    /**
     * Create a new <code>Picture</code>
     *
     * @param idx the index of the picture
     * @param parent the parent shape
     */
    public Picture(int idx, Shape parent) {
        base(null, parent);
        _escherContainer = CreateSpContainer(idx, parent is ShapeGroup);
    }

    /**
      * Create a <code>Picture</code> object
      *
      * @param escherRecord the <code>EscherSpContainer</code> record which holds information about
      *        this picture in the <code>Slide</code>
      * @param parent the parent shape of this picture
      */
     protected Picture(EscherContainerRecord escherRecord, Shape parent){
        base(escherRecord, parent);
    }

    /**
     * Returns index associated with this picture.
     * Index starts with 1 and points to a EscherBSE record which
     * holds information about this picture.
     *
     * @return the index to this picture (1 based).
     */
    public int GetPictureIndex(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherSimpleProperty prop = (EscherSimpleProperty)getEscherProperty(opt, EscherProperties.BLIP__BLIPTODISPLAY);
        return prop == null ? 0 : prop.GetPropertyValue();
    }

    /**
     * Create a new Picture and populate the Inital structure of the <code>EscherSp</code> record which holds information about this picture.

     * @param idx the index of the picture which referes to <code>EscherBSE</code> Container.
     * @return the create Picture object
     */
    protected EscherContainerRecord CreateSpContainer(int idx, bool IsChild) {
        _escherContainer = super.CreateSpContainer(isChild);
        _escherContainer.SetOptions((short)15);

        EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
        spRecord.SetOptions((short)((ShapeTypes.PictureFrame << 4) | 0x2));

        //set default properties for a picture
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        SetEscherProperty(opt, EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 0x800080);

        //another weird feature of powerpoint: for picture id we must add 0x4000.
        SetEscherProperty(opt, (short)(EscherProperties.BLIP__BLIPTODISPLAY + 0x4000), idx);

        return _escherContainer;
    }

    /**
     * Resize this picture to the default size.
     * For PNG and JPEG resizes the image to 100%,
     * for other types Sets the default size of 200x200 pixels.
     */
    public void SetDefaultSize(){
        PictureData pict = GetPictureData();
        if (pict  is Bitmap){
            BufferedImage img = null;
            try {
               	img = ImageIO.Read(new MemoryStream(pict.Data));
            }
            catch (IOException e){}
            catch (NegativeArraySizeException ne) {}

            if(img != null) {
                // Valid image, Set anchor from it
                SetAnchor(new java.awt.Rectangle(0, 0, img.Width*POINT_DPI/PIXEL_DPI, img.Height*POINT_DPI/PIXEL_DPI));
            } else {
                // Invalid image, go with the default metafile size
                SetAnchor(new java.awt.Rectangle(0, 0, 200, 200));
            }
        } else {
            //default size of a metafile picture is 200x200
            SetAnchor(new java.awt.Rectangle(50, 50, 200, 200));
        }
    }

    /**
     * Returns the picture data for this picture.
     *
     * @return the picture data for this picture.
     */
    public PictureData GetPictureData(){
        SlideShow ppt = Sheet.GetSlideShow();
        PictureData[] pict = ppt.GetPictureData();

        EscherBSERecord bse = GetEscherBSERecord();
        if (bse == null){
            logger.log(POILogger.ERROR, "no reference to picture data found ");
        } else {
            for ( int i = 0; i < pict.Length; i++ ) {
                if (pict[i].GetOffset() ==  bse.GetOffset()){
                    return pict[i];
                }
            }
            logger.log(POILogger.ERROR, "no picture found for our BSE offset " + bse.GetOffset());
        }
        return null;
    }

    protected EscherBSERecord GetEscherBSERecord(){
        SlideShow ppt = Sheet.GetSlideShow();
        Document doc = ppt.GetDocumentRecord();
        EscherContainerRecord dggContainer = doc.GetPPDrawingGroup().GetDggContainer();
        EscherContainerRecord bstore = (EscherContainerRecord)Shape.GetEscherChild(dggContainer, EscherContainerRecord.BSTORE_CONTAINER);
        if(bstore == null) {
            logger.log(POILogger.DEBUG, "EscherContainerRecord.BSTORE_CONTAINER was not found ");
            return null;
        }
        List lst = bstore.GetChildRecords();
        int idx = GetPictureIndex();
        if (idx == 0){
            logger.log(POILogger.DEBUG, "picture index was not found, returning ");
            return null;
        }
        return (EscherBSERecord)lst.Get(idx-1);
    }

    /**
     * Name of this picture.
     *
     * @return name of this picture
     */
    public String GetPictureName(){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        EscherComplexProperty prop = (EscherComplexProperty)getEscherProperty(opt, EscherProperties.BLIP__BLIPFILENAME);
        String name = null;
        if(prop != null){
            try {
                name = new String(prop.GetComplexData(), "UTF-16LE");
                int idx = name.IndexOf('\u0000');
                return idx == -1 ? name : name.Substring(0, idx);
            } catch (UnsupportedEncodingException e){
                throw new HSLFException(e);
            }
        }
        return name;
    }

    /**
     * Name of this picture.
     *
     * @param name of this picture
     */
    public void SetPictureName(String name){
        EscherOptRecord opt = (EscherOptRecord)getEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
        try {
            byte[] data = (name + '\u0000').GetBytes("UTF-16LE");
            EscherComplexProperty prop = new EscherComplexProperty(EscherProperties.BLIP__BLIPFILENAME, false, data);
            opt.AddEscherProperty(prop);
        } catch (UnsupportedEncodingException e){
            throw new HSLFException(e);
        }
    }

    /**
     * By default Set the orininal image size
     */
    protected void afterInsert(Sheet sh){
        super.afterInsert(sh);

        EscherBSERecord bse = GetEscherBSERecord();
        bse.SetRef(bse.GetRef() + 1);

        java.awt.Rectangle anchor = GetAnchor();
        if (anchor.Equals(new java.awt.Rectangle())){
            SetDefaultSize();
        }
    }

    public void Draw(Graphics2D graphics){
        AffineTransform at = graphics.GetTransform();
        ShapePainter.paint(this, graphics);

        PictureData data = GetPictureData();
        if(data != null) data.Draw(graphics, this);

        graphics.SetTransform(at);
    }
}





