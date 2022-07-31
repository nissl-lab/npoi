/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Text;
    using System.IO;
    using NPOI.DDF;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.SS.Util;
    using SixLabors.ImageSharp;


    /// <summary>
    /// Represents a escher picture.  Eg. A GIF, JPEG etc...
    /// @author Glen Stampoultzis
    /// @author Yegor Kozlov (yegor at apache.org)
    /// </summary>
    public class HSSFPicture : HSSFSimpleShape, IPicture
    {

        //int pictureIndex;
        //HSSFPatriarch patriarch;

        private static POILogger logger = POILogFactory.GetLogger(typeof(HSSFPicture));
        public HSSFPicture(EscherContainerRecord spContainer, ObjRecord objRecord)
            : base(spContainer, objRecord)
        {

        }
        /// <summary>
        /// Constructs a picture object.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <param name="anchor">The anchor.</param>
        public HSSFPicture(HSSFShape parent, HSSFAnchor anchor)
            : base(parent, anchor)
        {
            base.ShapeType = (OBJECT_TYPE_PICTURE);
            CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)GetObjRecord().SubRecords[0];
            cod.ObjectType = CommonObjectType.Picture;
        }
        protected override EscherContainerRecord CreateSpContainer()
        {
            EscherContainerRecord spContainer = base.CreateSpContainer();
            EscherOptRecord opt = (EscherOptRecord)spContainer.GetChildById(EscherOptRecord.RECORD_ID);
            opt.RemoveEscherProperty(EscherProperties.LINESTYLE__LINEDASHING);
            opt.RemoveEscherProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH);
            spContainer.RemoveChildRecord(spContainer.GetChildById(EscherTextboxRecord.RECORD_ID));
            return spContainer;
        }

        /// <summary>
        /// Reset the image to the dimension of the embedded image
        /// </summary>
        /// <remarks>
        /// Please note, that this method works correctly only for workbooks
        /// with default font size (Arial 10pt for .xls).
        /// If the default font is changed the resized image can be streched vertically or horizontally.
        /// </remarks>
        public void Resize()
        {
            Resize(Double.MaxValue);
        }

        /// <summary>
        /// Resize the image proportionally.
        /// </summary>
        /// <param name="scale">scale</param>
        /// <seealso cref="Resize(double, double)"/>
        public void Resize(double scale)
        {
            Resize(scale, scale);
        }

        /**
     * Resize the image
     * <p>
     * Please note, that this method works correctly only for workbooks
     * with default font size (Arial 10pt for .xls).
     * If the default font is changed the resized image can be streched vertically or horizontally.
     * </p>
     * <p>
     * <code>resize(1.0,1.0)</code> keeps the original size,<br/>
     * <code>resize(0.5,0.5)</code> resize to 50% of the original,<br/>
     * <code>resize(2.0,2.0)</code> resizes to 200% of the original.<br/>
     * <code>resize({@link Double#MAX_VALUE},{@link Double#MAX_VALUE})</code> resizes to the dimension of the embedded image. 
     * </p>
     *
     * @param scaleX the amount by which the image width is multiplied relative to the original width.
     * @param scaleY the amount by which the image height is multiplied relative to the original height.
     */
        public void Resize(double scaleX, double scaleY)
        {
            HSSFClientAnchor anchor = (HSSFClientAnchor)ClientAnchor;
            anchor.AnchorType = AnchorType.MoveDontResize;

            HSSFClientAnchor pref = GetPreferredSize(scaleX, scaleY) as HSSFClientAnchor;

            int row2 = anchor.Row1 + (pref.Row2 - pref.Row1);
            int col2 = anchor.Col1 + (pref.Col2 - pref.Col1);

            anchor.Col2=((short)col2);
            // anchor.setDx1(0);
            anchor.Dx2=(pref.Dx2);

            anchor.Row2 = (row2);
            // anchor.setDy1(0);
            anchor.Dy2 = (pref.Dy2);
        }
        /// <summary>
        /// Gets or sets the index of the picture.
        /// </summary>
        /// <value>The index of the picture.</value>
        public int PictureIndex
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.BLIP__BLIPTODISPLAY);
                if (null == property)
                {
                    return -1;
                }
                return property.PropertyValue;
            }
            set
            {
                SetPropertyValue(new EscherSimpleProperty(EscherProperties.BLIP__BLIPTODISPLAY, false, true, value));
            }
        }
        /**
         * Calculate the preferred size for this picture.
         *
         * @param scale the amount by which image dimensions are multiplied relative to the original size.
         * @return HSSFClientAnchor with the preferred size for this image
         * @since POI 3.0.2
         */
        public IClientAnchor GetPreferredSize(double scale)
        {
            return GetPreferredSize(scale, scale);
        }
        /// <summary>
        /// Calculate the preferred size for this picture.
        /// </summary>
        /// <param name="scaleX">the amount by which image width is multiplied relative to the original width.</param>
        /// <param name="scaleY">the amount by which image height is multiplied relative to the original height.</param>
        /// <returns>HSSFClientAnchor with the preferred size for this image</returns>
        public IClientAnchor GetPreferredSize(double scaleX, double scaleY)
        {
            ImageUtils.SetPreferredSize(this, scaleX, scaleY);
            return ClientAnchor;
            
        }

        /// <summary>
        /// Calculate the preferred size for this picture.
        /// </summary>
        /// <returns>HSSFClientAnchor with the preferred size for this image</returns>
        public IClientAnchor GetPreferredSize()
        {
            return GetPreferredSize(1.0);
        }


        /// <summary>
        /// The metadata of PNG and JPEG can contain the width of a pixel in millimeters.
        /// Return the the "effective" dpi calculated as 
        /// <c>25.4/HorizontalPixelSize</c>
        /// and 
        /// <c>25.4/VerticalPixelSize</c>
        /// .  Where 25.4 is the number of mm in inch.
        /// </summary>
        /// <param name="r">The image.</param>
        /// <returns>the resolution</returns>
        protected Size GetResolution(Image r)
        {
            //int hdpi = 96, vdpi = 96;
            //double mm2inch = 25.4;
            return new Size((int)r.Metadata.HorizontalResolution, (int)r.Metadata.VerticalResolution);
        }

        /// <summary>
        /// Return the dimension of the embedded image in pixel
        /// </summary>
        /// <returns>image dimension</returns>
        public Size GetImageDimension()
        {
            InternalWorkbook iwb = (_patriarch.Sheet.Workbook as HSSFWorkbook).Workbook;
            EscherBSERecord bse = iwb.GetBSERecord(PictureIndex);
            byte[] data = bse.BlipRecord.PictureData;
            //int type = bse.BlipTypeWin32;

            using (MemoryStream ms = RecyclableMemory.GetStream(data))
            {
                using (Image img = Image.Load(ms))
                {
                    return img.Size();
                }
            }
        }
        /**
         * Return picture data for this shape
         *
         * @return picture data for this shape
         */
        public IPictureData PictureData
        {
            get
            {
                InternalWorkbook iwb = ((_patriarch.Sheet.Workbook) as HSSFWorkbook).Workbook;
                EscherBSERecord bse = iwb.GetBSERecord(PictureIndex);
                EscherBlipRecord blipRecord = bse.BlipRecord;
                return new HSSFPictureData(blipRecord);
            }
        }


        internal override void AfterInsert(HSSFPatriarch patriarch)
        {
            EscherAggregate agg = patriarch.GetBoundAggregate();
            agg.AssociateShapeToObjRecord(GetEscherContainer().GetChildById(EscherClientDataRecord.RECORD_ID), GetObjRecord());
            if (PictureIndex != -1)
            {
                EscherBSERecord bse =
                    (patriarch.Sheet.Workbook as HSSFWorkbook).Workbook.GetBSERecord(PictureIndex);
                bse.Ref = (bse.Ref + 1);
            }
        }

        /**
         * The color applied to the lines of this shape.
         */
        public String FileName
        {
            get
            {
                EscherComplexProperty propFile = (EscherComplexProperty)GetOptRecord().Lookup(
                              EscherProperties.BLIP__BLIPFILENAME);
                return (null == propFile) ? "" : Trim(StringUtil.GetFromUnicodeLE(propFile.ComplexData));
            }
            set
            {
                // TODO: add trailing \u0000? 
                byte[] bytes = StringUtil.GetToUnicodeLE(value);
                EscherComplexProperty prop = new EscherComplexProperty(EscherProperties.BLIP__BLIPFILENAME, true, bytes);
                SetPropertyValue(prop);
            }
        }
        private String Trim(string value)
        {
            int end = value.Length;
            int st = 0;
            //int off = offset;      /* avoid getfield opcode */
            char[] val = value.ToCharArray();    /* avoid getfield opcode */

            while ((st < end) && (val[st] <= ' '))
            {
                st++;
            }
            while ((st < end) && (val[end - 1] <= ' '))
            {
                end--;
            }
            return ((st > 0) || (end < value.Length)) ? value.Substring(st, end - st) : value;
        }

        public override int ShapeType
        {
            get { return base.ShapeType; }
            set
            {
                throw new InvalidOperationException("Shape type can not be changed in " + this.GetType().Name);
            }
        }


        internal override HSSFShape CloneShape()
        {
            EscherContainerRecord spContainer = new EscherContainerRecord();
            byte[] inSp = GetEscherContainer().Serialize();
            spContainer.FillFields(inSp, 0, new DefaultEscherRecordFactory());
            ObjRecord obj = (ObjRecord)GetObjRecord().CloneViaReserialise();
            return new HSSFPicture(spContainer, obj);
        }


        /**
         * @return the anchor that is used by this picture.
         */
        public IClientAnchor ClientAnchor
        {
            get
            {
                HSSFAnchor a = Anchor;
                return (a is HSSFClientAnchor) ? (HSSFClientAnchor)a : null;
            }
        }


        /**
         * @return the sheet which contains the picture shape
         */
        public ISheet Sheet
        {
            get
            {
                return Patriarch.Sheet;
            }
        }
    }
}