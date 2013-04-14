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
    using System.Drawing;
    using System.Text;
    using System.IO;
    using NPOI.DDF;
    using NPOI.Util;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;


    /// <summary>
    /// Represents a escher picture.  Eg. A GIF, JPEG etc...
    /// @author Glen Stampoultzis
    /// @author Yegor Kozlov (yegor at apache.org)
    /// </summary>
    public class HSSFPicture : HSSFSimpleShape, IPicture
    {
        /**
         * width of 1px in columns with default width in Units of 1/256 of a Char width
         */
        private const float PX_DEFAULT = 32.00f;
        /**
         * width of 1px in columns with overridden width in Units of 1/256 of a Char width
         */
        private const float PX_MODIFIED = 36.56f;

        /**
         * Height of 1px of a row
         */
        private const int PX_ROW = 15;

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

        /// <summary>
        /// Reset the image to the original size.
        /// </summary>
        public void Resize(double scale)
        {
            HSSFClientAnchor anchor = (HSSFClientAnchor)Anchor;
            anchor.AnchorType = 2;

            IClientAnchor pref = GetPreferredSize(scale);

            int row2 = anchor.Row1 + (pref.Row2 - pref.Row1);
            int col2 = anchor.Col1 + (pref.Col2 - pref.Col1);

            anchor.Col2 = col2;
            anchor.Dx1 = 0;
            anchor.Dx2 = pref.Dx2;

            anchor.Row2 = row2;
            anchor.Dy1 = 0;
            anchor.Dy2 = pref.Dy2;
        }
        /// <summary>
        /// Reset the image to the original size.
        /// </summary>
        public void Resize()
        {
            Resize(1.0);
        }
        /// <summary>
        /// Calculate the preferred size for this picture.
        /// </summary>
        /// <param name="scale">the amount by which image dimensions are multiplied relative to the original size.</param>
        /// <returns>HSSFClientAnchor with the preferred size for this image</returns>
        public HSSFClientAnchor GetPreferredSize(double scale)
        {
            HSSFClientAnchor anchor = (HSSFClientAnchor)Anchor;

            Size size = GetImageDimension();
            double scaledWidth = size.Width * scale;
            double scaledHeight = size.Height * scale;

            float w = 0;

            //space in the leftmost cell
            w += GetColumnWidthInPixels(anchor.Col1) * (1 - (float)anchor.Dx1 / 1024);
            short col2 = (short)(anchor.Col1 + 1);
            int dx2 = 0;

            while (w < scaledWidth)
            {
                w += GetColumnWidthInPixels(col2++);
            }

            if (w > scaledWidth)
            {
                //calculate dx2, offset in the rightmost cell
                col2--;
                double cw = GetColumnWidthInPixels(col2);
                double delta = w - scaledWidth;
                dx2 = (int)((cw - delta) / cw * 1024);
            }
            anchor.Col2 = col2;
            anchor.Dx2 = dx2;

            float h = 0;
            h += (1 - (float)anchor.Dy1 / 256) * GetRowHeightInPixels(anchor.Row1);
            int row2 = anchor.Row1 + 1;
            int dy2 = 0;

            while (h < scaledHeight)
            {
                h += GetRowHeightInPixels(row2++);
            }
            if (h > scaledHeight)
            {
                row2--;
                double ch = GetRowHeightInPixels(row2);
                double delta = h - scaledHeight;
                dy2 = (int)((ch - delta) / ch * 256);
            }
            anchor.Row2 = row2;
            anchor.Dy2 = dy2;

            return anchor;
        }

        /// <summary>
        /// Calculate the preferred size for this picture.
        /// </summary>
        /// <returns>HSSFClientAnchor with the preferred size for this image</returns>
        public NPOI.SS.UserModel.IClientAnchor GetPreferredSize()
        {
            return GetPreferredSize(1.0);
        }

        /// <summary>
        /// Gets the column width in pixels.
        /// </summary>
        /// <param name="column">The column.</param>
        /// <returns></returns>
        private float GetColumnWidthInPixels(int column)
        {

            int cw = _patriarch.Sheet.GetColumnWidth(column);
            float px = GetPixelWidth(column);

            return cw / px;
        }

        /// <summary>
        /// Gets the row height in pixels.
        /// </summary>
        /// <param name="i">The row</param>
        /// <returns></returns>
        private float GetRowHeightInPixels(int i)
        {

            IRow row = _patriarch.Sheet.GetRow(i);
            float height;
            if (row != null) height = row.Height;
            else height = _patriarch.Sheet.DefaultRowHeight;

            return height / PX_ROW;
        }

        /// <summary>
        /// Gets the width of the pixel.
        /// </summary>
        /// <param name="column">The column.</param>
        /// <returns></returns>
        private float GetPixelWidth(int column)
        {

            int def = _patriarch.Sheet.DefaultColumnWidth * 256;
            int cw = _patriarch.Sheet.GetColumnWidth(column);

            return cw == def ? PX_DEFAULT : PX_MODIFIED;
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

            //NodeList lst;
            //Element node = (Element)r.GetImageMetadata(0).GetAsTree("javax_imageio_1.0");
            //lst = node.GetElementsByTagName("HorizontalPixelSize");
            //if (lst != null && lst.GetLength == 1) hdpi = (int)(mm2inch / Float.ParseFloat(((Element)lst.item(0)).GetAttribute("value")));

            //lst = node.GetElementsByTagName("VerticalPixelSize");
            //if (lst != null && lst.GetLength == 1) vdpi = (int)(mm2inch / Float.ParseFloat(((Element)lst.item(0)).GetAttribute("value")));

            return new Size((int)r.HorizontalResolution, (int)r.VerticalResolution);
        }

        /// <summary>
        /// Return the dimension of this image
        /// </summary>
        /// <returns>image dimension</returns>
        public Size GetImageDimension()
        {
            EscherBSERecord bse = (_patriarch.Sheet.Workbook as HSSFWorkbook).Workbook.GetBSERecord(PictureIndex);
            byte[] data = bse.BlipRecord.PictureData;
            //int type = bse.BlipTypeWin32;

            using (MemoryStream ms = new MemoryStream(data))
            {
                using (Image img = Image.FromStream(ms))
                {
                    return img.Size;
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
                EscherBlipRecord blipRecord = iwb.GetBSERecord(PictureIndex).BlipRecord;
                return new HSSFPictureData(blipRecord);
            }
        }


        internal override void AfterInsert(HSSFPatriarch patriarch)
        {
            EscherAggregate agg = patriarch.GetBoundAggregate();
            agg.AssociateShapeToObjRecord(GetEscherContainer().GetChildById(EscherClientDataRecord.RECORD_ID), GetObjRecord());
            EscherBSERecord bse =
                    (patriarch.Sheet.Workbook as HSSFWorkbook).Workbook.GetBSERecord(PictureIndex);
            bse.Ref = (bse.Ref + 1);
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
                try
                {
                    if (null == propFile)
                    {
                        return "";
                    }
                    return Trim(Encoding.Unicode.GetString(propFile.ComplexData));
                }
                catch (Exception)
                {
                    return "";
                }
            }
            set
            {
                try
                {
                    EscherComplexProperty prop = new EscherComplexProperty(EscherProperties.BLIP__BLIPFILENAME, true,
                        Encoding.Unicode.GetBytes(value));
                    SetPropertyValue(prop);
                }
                catch (Exception)
                {
                    logger.Log(POILogger.ERROR, "Unsupported encoding: UTF-16LE");
                }
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
    }
}