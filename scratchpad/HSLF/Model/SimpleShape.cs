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

using NPOI.DDF;
using NPOI.Util;
using System.Drawing;
using NPOI.HSLF.Record;
using System.IO;
using NPOI.HSLF.Exceptions;
using System;
namespace NPOI.HSLF.Model
{

    /**
     *  An abstract simple (non-group) shape.
     *  This is the parent class for all primitive shapes like Line, Rectangle, etc.
     *
     *  @author Yegor Kozlov
     */
    public abstract class SimpleShape : Shape
    {

        public static double DEFAULT_LINE_WIDTH = 0.75;

        /**
         * Records stored in EscherClientDataRecord
         */
        protected Record[] _clientRecords;
        protected EscherClientDataRecord _clientData;

        /**
         * Create a SimpleShape object and Initialize it from the supplied Record Container.
         *
         * @param escherRecord    <code>EscherSpContainer</code> Container which holds information about this shape
         * @param parent    the parent of the shape
         */
        protected SimpleShape(EscherContainerRecord escherRecord, Shape parent)
            : base(escherRecord, parent)
        {

        }

        /**
         * Create a new Shape
         *
         * @param isChild   <code>true</code> if the Line is inside a group, <code>false</code> otherwise
         * @return the record Container which holds this shape
         */
        protected EscherContainerRecord CreateSpContainer(bool IsChild)
        {
            _escherContainer = new EscherContainerRecord();
            _escherContainer.SetRecordId(EscherContainerRecord.SP_CONTAINER);
            _escherContainer.SetOptions((short)15);

            EscherSpRecord sp = new EscherSpRecord();
            int flags = EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE;
            if (isChild) flags |= EscherSpRecord.FLAG_CHILD;
            sp.SetFlags(flags);
            _escherContainer.AddChildRecord(sp);

            EscherOptRecord opt = new EscherOptRecord();
            opt.SetRecordId(EscherOptRecord.RECORD_ID);
            _escherContainer.AddChildRecord(opt);

            EscherRecord anchor;
            if (isChild) anchor = new EscherChildAnchorRecord();
            else
            {
                anchor = new EscherClientAnchorRecord();

                //hack. internal variable EscherClientAnchorRecord.shortRecord can be
                //Initialized only in FillFields(). We need to Set shortRecord=false;
                byte[] header = new byte[16];
                LittleEndian.PutUshort(header, 0, 0);
                LittleEndian.PutUshort(header, 2, 0);
                LittleEndian.PutInt(header, 4, 8);
                anchor.FillFields(header, 0, null);
            }
            _escherContainer.AddChildRecord(anchor);

            return _escherContainer;
        }

        /**
         *  Returns width of the line in in points
         */
        public double GetLineWidth()
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            EscherSimpleProperty prop = (EscherSimpleProperty)GetEscherProperty(opt, EscherProperties.LINESTYLE__LINEWIDTH);
            double width = prop == null ? DEFAULT_LINE_WIDTH : (double)prop.PropertyValue / EMU_PER_POINT;
            return width;
        }

        /**
         *  Sets the width of line in in points
         *  @param width  the width of line in in points
         */
        public void SetLineWidth(double width)
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            SetEscherProperty(opt, EscherProperties.LINESTYLE__LINEWIDTH, (int)(width * EMU_PER_POINT));
        }

        /**
         * Sets the color of line
         *
         * @param color new color of the line
         */
        public void SetLineColor(Color color)
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            if (color == null)
            {
                SetEscherProperty(opt, EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x80000);
            }
            else
            {
                int rgb = new Color(color.Blue, color.Green, color.Red, 0).GetRGB();
                SetEscherProperty(opt, EscherProperties.LINESTYLE__COLOR, rgb);
                SetEscherProperty(opt, EscherProperties.LINESTYLE__NOLINEDRAWDASH, color == null ? 0x180010 : 0x180018);
            }
        }

        /**
         * @return color of the line. If color is not Set returns <code>java.awt.Color.black</code>
         */
        public Color GetLineColor()
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);

            EscherSimpleProperty p1 = (EscherSimpleProperty)GetEscherProperty(opt, EscherProperties.LINESTYLE__COLOR);
            EscherSimpleProperty p2 = (EscherSimpleProperty)GetEscherProperty(opt, EscherProperties.LINESTYLE__NOLINEDRAWDASH);
            int p2val = p2 == null ? 0 : p2.PropertyValue;
            Color clr = null;
            if ((p2val & 0x8) != 0 || (p2val & 0x10) != 0)
            {
                int rgb = p1 == null ? 0 : p1.PropertyValue;
                if (rgb >= 0x8000000)
                {
                    int idx = rgb % 0x8000000;
                    if (Sheet != null)
                    {
                        ColorSchemeAtom ca = Sheet.GetColorScheme();
                        if (idx >= 0 && idx <= 7) rgb = ca.GetColor(idx);
                    }
                }
                Color tmp = new Color(rgb, true);
                clr = new Color(tmp.GetBlue(), tmp.GetGreen(), tmp.GetRed());
            }
            return clr;
        }

        /**
         * Gets line dashing. One of the PEN_* constants defined in this class.
         *
         * @return dashing of the line.
         */
        public int GetLineDashing()
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);

            EscherSimpleProperty prop = (EscherSimpleProperty)GetEscherProperty(opt, EscherProperties.LINESTYLE__LINEDASHING);
            return prop == null ? Line.PEN_SOLID : prop.PropertyValue;
        }

        /**
         * Sets line dashing. One of the PEN_* constants defined in this class.
         *
         * @param pen new style of the line.
         */
        public void SetLineDashing(int pen)
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);

            SetEscherProperty(opt, EscherProperties.LINESTYLE__LINEDASHING, pen == Line.PEN_SOLID ? -1 : pen);
        }

        /**
         * Sets line style. One of the constants defined in this class.
         *
         * @param style new style of the line.
         */
        public void SetLineStyle(int style)
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            SetEscherProperty(opt, EscherProperties.LINESTYLE__LINESTYLE, style == Line.LINE_SIMPLE ? -1 : style);
        }

        /**
         * Returns line style. One of the constants defined in this class.
         *
         * @return style of the line.
         */
        public int GetLineStyle()
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            EscherSimpleProperty prop = (EscherSimpleProperty)GetEscherProperty(opt, EscherProperties.LINESTYLE__LINESTYLE);
            return prop == null ? Line.LINE_SIMPLE : prop.PropertyValue;
        }

        /**
         * The color used to fill this shape.
         */
        public Color GetFillColor()
        {
            return GetFill().GetForegroundColor();
        }

        /**
         * The color used to fill this shape.
         *
         * @param color the background color
         */
        public void SetFillColor(Color color)
        {
            GetFill().SetForegroundColor(color);
        }

        /**
         * Whether the shape is horizontally flipped
         *
         * @return whether the shape is horizontally flipped
         */
        public bool GetFlipHorizontal()
        {
            EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            return (spRecord.Flags & EscherSpRecord.FLAG_FLIPHORIZ) != 0;
        }

        /**
         * Whether the shape is vertically flipped
         *
         * @return whether the shape is vertically flipped
         */
        public bool GetFlipVertical()
        {
            EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            return (spRecord.Flags & EscherSpRecord.FLAG_FLIPVERT) != 0;
        }

        /**
         * Rotation angle in degrees
         *
         * @return rotation angle in degrees
         */
        public int GetRotation()
        {
            int rot = GetEscherProperty(EscherProperties.TRANSFORM__ROTATION);
            int angle = (rot >> 16) % 360;

            return angle;
        }

        /**
         * Rotate this shape
         *
         * @param theta the rotation angle in degrees
         */
        public void SetRotation(int theta)
        {
            SetEscherProperty(EscherProperties.TRANSFORM__ROTATION, (theta << 16));
        }

        public Rectangle2D GetLogicalAnchor2D()
        {
            Rectangle anchor = GetAnchor2D();

            //if it is a groupped shape see if we need to transform the coordinates
            if (_parent != null)
            {
                Shape top = _parent;
                while (top.GetParent() != null) top = top.GetParent();

                Rectangle clientAnchor = top.GetAnchor2D();
                Rectangle spgrAnchor = ((ShapeGroup)top).GetCoordinates();

                double scalex = spgrAnchor.Width / clientAnchor.Width;
                double scaley = spgrAnchor.Height / clientAnchor.Height;

                double x = clientAnchor.X + (anchor.X - spgrAnchor.X) / scalex;
                double y = clientAnchor.Y + (anchor.Y - spgrAnchor.Y) / scaley;
                double width = anchor.Width / scalex;
                double height = anchor.Height / scaley;

                anchor = new Rectangle2D.Double(x, y, width, height);

            }

            int angle = GetRotation();
            if (angle != 0)
            {
                double centerX = anchor.X + anchor.Width / 2;
                double centerY = anchor.Y + anchor.Height / 2;

                AffineTransform trans = new AffineTransform();
                trans.translate(centerX, centerY);
                trans.rotate(Math.ToRadians(angle));
                trans.translate(-centerX, -centerY);

                Rectangle2D rect = trans.CreateTransformedShape(anchor).GetBounds2D();
                if ((anchor.Width < anchor.Height && rect.Width > rect.Height) ||
                    (anchor.Width > anchor.Height && rect.Width < rect.Height))
                {
                    trans = new AffineTransform();
                    trans.translate(centerX, centerY);
                    trans.rotate(Math.PI / 2);
                    trans.translate(-centerX, -centerY);
                    anchor = trans.CreateTransformedShape(anchor).GetBounds2D();
                }
            }
            return anchor;
        }

        public void Draw(Graphics2D graphics)
        {
            AffineTransform at = graphics.GetTransform();
            ShapePainter.paint(this, graphics);
            graphics.SetTransform(at);
        }

        /**
         *  Find a record in the underlying EscherClientDataRecord
         *
         * @param recordType type of the record to search
         */
        protected Record GetClientDataRecord(int recordType)
        {

            Record[] records = GetClientRecords();
            if (records != null) for (int i = 0; i < records.Length; i++)
                {
                    if (records[i].GetRecordType() == recordType)
                    {
                        return records[i];
                    }
                }
            return null;
        }

        /**
         * Search for EscherClientDataRecord, if found, convert its contents into an array of HSLF records
         *
         * @return an array of HSLF records Contained in the shape's EscherClientDataRecord or <code>null</code>
         */
        protected Record[] GetClientRecords()
        {
            if (_clientData == null)
            {
                EscherRecord r = Shape.GetEscherChild(GetSpContainer(), EscherClientDataRecord.RECORD_ID);
                //ddf can return EscherContainerRecord with recordId=EscherClientDataRecord.RECORD_ID
                //convert in to EscherClientDataRecord on the fly
                if (r != null && !(r is EscherClientDataRecord))
                {
                    byte[] data = r.Serialize();
                    r = new EscherClientDataRecord();
                    r.FillFields(data, 0, new DefaultEscherRecordFactory());
                }
                _clientData = (EscherClientDataRecord)r;
            }
            if (_clientData != null && _clientRecords == null)
            {
                byte[] data = _clientData.GetRemainingData();
                _clientRecords = Record.FindChildRecords(data, 0, data.Length);
            }
            return _clientRecords;
        }

        protected void UpdateClientData()
        {
            if (_clientData != null && _clientRecords != null)
            {
                MemoryStream out1 = new MemoryStream();
                try
                {
                    for (int i = 0; i < _clientRecords.Length; i++)
                    {
                        _clientRecords[i].WriteOut(out1);
                    }
                }
                catch (Exception e)
                {
                    throw new HSLFException(e);
                }
                _clientData.SetRemainingData(out1.ToArray());
            }
        }

        public void SetHyperlink(Hyperlink link)
        {
            if (link.GetId() == -1)
            {
                throw new HSLFException("You must call SlideShow.AddHyperlink(Hyperlink link) first");
            }

            EscherClientDataRecord cldata = new EscherClientDataRecord();
            cldata.SetOptions((short)0xF);
            GetSpContainer().AddChildRecord(cldata); // TODO - junit to prove GetChildRecords().add is wrong

            InteractiveInfo info = new InteractiveInfo();
            InteractiveInfoAtom infoAtom = info.GetInteractiveInfoAtom();

            switch (link.GetType())
            {
                case Hyperlink.LINK_FIRSTSLIDE:
                    infoAtom.SetAction(InteractiveInfoAtom.ACTION_JUMP);
                    infoAtom.SetJump(InteractiveInfoAtom.JUMP_FIRSTSLIDE);
                    infoAtom.SetHyperlinkType(InteractiveInfoAtom.LINK_FirstSlide);
                    break;
                case Hyperlink.LINK_LASTSLIDE:
                    infoAtom.SetAction(InteractiveInfoAtom.ACTION_JUMP);
                    infoAtom.SetJump(InteractiveInfoAtom.JUMP_LASTSLIDE);
                    infoAtom.SetHyperlinkType(InteractiveInfoAtom.LINK_LastSlide);
                    break;
                case Hyperlink.LINK_NEXTSLIDE:
                    infoAtom.SetAction(InteractiveInfoAtom.ACTION_JUMP);
                    infoAtom.SetJump(InteractiveInfoAtom.JUMP_NEXTSLIDE);
                    infoAtom.SetHyperlinkType(InteractiveInfoAtom.LINK_NextSlide);
                    break;
                case Hyperlink.LINK_PREVIOUSSLIDE:
                    infoAtom.SetAction(InteractiveInfoAtom.ACTION_JUMP);
                    infoAtom.SetJump(InteractiveInfoAtom.JUMP_PREVIOUSSLIDE);
                    infoAtom.SetHyperlinkType(InteractiveInfoAtom.LINK_PreviousSlide);
                    break;
                case Hyperlink.LINK_URL:
                    infoAtom.SetAction(InteractiveInfoAtom.ACTION_HYPERLINK);
                    infoAtom.SetJump(InteractiveInfoAtom.JUMP_NONE);
                    infoAtom.SetHyperlinkType(InteractiveInfoAtom.LINK_Url);
                    break;
            }

            infoAtom.SetHyperlinkID(link.GetId());

            MemoryStream out1 = new MemoryStream();
            try
            {
                info.WriteOut(out1);
            }
            catch (Exception e)
            {
                throw new HSLFException(e);
            }
            cldata.SetRemainingData(out1.ToArray());

        }

    }
}





