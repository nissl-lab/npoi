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

namespace NPOI.HSLF.Model
{

    using NPOI.DDF;
    using NPOI.HSLF.Record;
    using NPOI.Util;
    using System;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;
    using System.Drawing;

    /**
     *  <p>
      * Represents a Shape which is the elemental object that composes a Drawing.
     *  This class is a wrapper around EscherSpContainer which holds all information
     *  about a shape in PowerPoint document.
     *  </p>
     *  <p>
     *  When you add a shape, you usually specify the dimensions of the shape and the position
     *  of the upper'left corner of the bounding box for the shape relative to the upper'left
     *  corner of the page, worksheet, or slide. Distances in the Drawing layer are measured
     *  in points (72 points = 1 inch).
     *  </p>
     * <p>
      *
      * @author Yegor Kozlov
     */
    public abstract class Shape
    {

        // For logging
        protected POILogger logger = POILogFactory.GetLogger(typeof(Shape));

        /**
         * In Escher absolute distances are specified in
         * English Metric Units (EMUs), occasionally referred to as A units;
         * there are 360000 EMUs per centimeter, 914400 EMUs per inch, 12700 EMUs per point.
         */
        public static int EMU_PER_INCH = 914400;
        public static int EMU_PER_POINT = 12700;
        public static int EMU_PER_CENTIMETER = 360000;

        /**
         * Master DPI (576 pixels per inch).
         * Used by the reference coordinate system in PowerPoint.
         */
        public static int MASTER_DPI = 576;

        /**
         * Pixels DPI (96 pixels per inch)
         */
        public static int PIXEL_DPI = 96;

        /**
         * Points DPI (72 pixels per inch)
         */
        public static int POINT_DPI = 72;

        /**
         * Either EscherSpContainer or EscheSpgrContainer record
         * which holds information about this shape.
         */
        protected EscherContainerRecord _escherContainer;

        /**
         * Parent of this shape.
         * <code>null</code> for the topmost shapes.
         */
        protected Shape _parent;

        /**
         * The <code>Sheet</code> this shape belongs to
         */
        protected Sheet _sheet;

        /**
         * Fill
         */
        protected Fill _Fill;

        /**
         * Create a Shape object. This constructor is used when an existing Shape is read from from a PowerPoint document.
         *
         * @param escherRecord       <code>EscherSpContainer</code> Container which holds information about this shape
         * @param parent             the parent of this Shape
         */
        protected Shape(EscherContainerRecord escherRecord, Shape parent)
        {
            _escherContainer = escherRecord;
            _parent = parent;
        }

        /**
         * Creates the lowerlevel escher records for this shape.
         */
        protected abstract EscherContainerRecord CreateSpContainer(bool IsChild);

        /**
         *  @return the parent of this shape
         */
        public Shape GetParent()
        {
            return _parent;
        }

        /**
         * @return name of the shape.
         */
        public String GetShapeName()
        {
            return Enum.GetName(typeof(ShapeTypes),GetShapeType());
        }

        /**
         * @return type of the shape.
         * @see NPOI.HSLF.record.RecordTypes
         */
        public int GetShapeType()
        {
            EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            return spRecord.Options >> 4;
        }

        /**
         * @param type type of the shape.
         * @see NPOI.HSLF.record.RecordTypes
         */
        public void SetShapeType(int type)
        {
            EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            spRecord.Options = (short)(type << 4 | 0x2);
        }

        /**
         * Returns the anchor (the bounding box rectangle) of this shape.
         * All coordinates are expressed in points (72 dpi).
         *
         * @return the anchor of this shape
         */
        //public Rectangle GetAnchor()
        //{
        //    Rectangle anchor2d = GetAnchor2D();
        //    return anchor2d.Bounds;
        //}

        /**
         * Returns the anchor (the bounding box rectangle) of this shape.
         * All coordinates are expressed in points (72 dpi).
         *
         * @return the anchor of this shape
         */
        public Rectangle GetAnchor2D()
        {
            EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            int flags = spRecord.Flags;
            Rectangle anchor;
            if ((flags & EscherSpRecord.FLAG_CHILD) != 0)
            {
                EscherChildAnchorRecord rec = (EscherChildAnchorRecord)GetEscherChild(_escherContainer, EscherChildAnchorRecord.RECORD_ID);
                anchor = new Rectangle();
                if (rec == null)
                {
                    logger.Log(POILogger.WARN, "EscherSpRecord.FLAG_CHILD is Set but EscherChildAnchorRecord was not found");
                    EscherClientAnchorRecord clrec = (EscherClientAnchorRecord)GetEscherChild(_escherContainer, EscherClientAnchorRecord.RECORD_ID);
                    anchor = new Rectangle();
                    anchor.X = clrec.Col1 * POINT_DPI / MASTER_DPI;
                    anchor.Y= clrec.Flag * POINT_DPI / MASTER_DPI;
                    anchor.Width = (clrec.Dx1 - clrec.Col1) * POINT_DPI / MASTER_DPI;
                    anchor.Height = (clrec.Row1 - clrec.Flag) * POINT_DPI / MASTER_DPI;
                }
                else
                {
                    anchor.X = rec.Dx1 * POINT_DPI / MASTER_DPI;
                    anchor.Y = rec.Dy1 * POINT_DPI / MASTER_DPI;
                    anchor.Width = (rec.Dx2 - rec.Dx1) * POINT_DPI / MASTER_DPI;
                    anchor.Height = (rec.Dy2 - rec.Dy1) * POINT_DPI / MASTER_DPI;
                }
            }
            else
            {
                EscherClientAnchorRecord rec = (EscherClientAnchorRecord)GetEscherChild(_escherContainer, EscherClientAnchorRecord.RECORD_ID);
                anchor = new Rectangle();
                anchor.X = rec.Col1 * POINT_DPI / MASTER_DPI;
                anchor.Y = rec.Flag * POINT_DPI / MASTER_DPI;
                anchor.Width = (rec.Dx1 - rec.Col1) * POINT_DPI / MASTER_DPI;
                anchor.Height = (rec.Row1 - rec.Flag) * POINT_DPI / MASTER_DPI;
            }
            return anchor;
        }

        public Rectangle GetLogicalAnchor2D()
        {
            return GetAnchor2D();
        }

        /**
         * Sets the anchor (the bounding box rectangle) of this shape.
         * All coordinates should be expressed in points (72 dpi).
         *
         * @param anchor new anchor
         */
        public void SetAnchor(Rectangle anchor)
        {
            EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            int flags = spRecord.Flags;
            if ((flags & EscherSpRecord.FLAG_CHILD) != 0)
            {
                EscherChildAnchorRecord rec = (EscherChildAnchorRecord)GetEscherChild(_escherContainer, EscherChildAnchorRecord.RECORD_ID);
                rec.Dx1 = ((int)(anchor.X * MASTER_DPI / POINT_DPI));
                rec.Dy1=((int)(anchor.Y * MASTER_DPI / POINT_DPI));
                rec.Dx2=((int)((anchor.Width + anchor.X) * MASTER_DPI / POINT_DPI));
                rec.Dy2=((int)((anchor.Height + anchor.Y) * MASTER_DPI / POINT_DPI));
            }
            else
            {
                EscherClientAnchorRecord rec = (EscherClientAnchorRecord)GetEscherChild(_escherContainer, EscherClientAnchorRecord.RECORD_ID);
                rec.Flag=((short)(anchor.Y * MASTER_DPI / POINT_DPI));
                rec.Col1=((short)(anchor.X * MASTER_DPI / POINT_DPI));
                rec.Dx1 = ((short)(((anchor.Width + anchor.X) * MASTER_DPI / POINT_DPI)));
                rec.Row1=((short)(((anchor.Height + anchor.Y) * MASTER_DPI / POINT_DPI)));
            }

        }

        /**
         * Moves the top left corner of the shape to the specified point.
         *
         * @param x the x coordinate of the top left corner of the shape
         * @param y the y coordinate of the top left corner of the shape
         */
        public void MoveTo(float x, float y)
        {
            Rectangle anchor = GetAnchor2D();
            anchor.X=Convert.ToInt32(x);
            anchor.Y=Convert.ToInt32(y);
            anchor.Width=anchor.Width;
            anchor.Height=anchor.Height;
            SetAnchor(anchor);
        }

        /**
         * Helper method to return escher child by record ID
         *
         * @return escher record or <code>null</code> if not found.
         */
        public static EscherRecord GetEscherChild(EscherContainerRecord owner, int recordId)
        {
            for (List<EscherRecord>.Enumerator iterator = owner.GetChildIterator(); iterator.MoveNext(); )
            {
                EscherRecord escherRecord = iterator.Current;
                if (escherRecord.RecordId == recordId)
                    return escherRecord;
            }
            return null;
        }

        /**
         * Returns  escher property by id.
         *
         * @return escher property or <code>null</code> if not found.
         */
        public static EscherProperty GetEscherProperty(EscherOptRecord opt, int propId)
        {
            if (opt != null)
            {
                for (List<EscherProperty>.Enumerator iterator = opt.EscherProperties.GetEnumerator(); iterator.MoveNext(); )
                {
                    EscherProperty prop = (EscherProperty)iterator.Current;
                    if (prop.PropertyNumber == propId)
                        return prop;
                }
            }
            return null;
        }

        /**
         * Set an escher property for this shape.
         *
         * @param opt       The opt record to Set the properties to.
         * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
         * @param value     value of the property. If value = -1 then the property is Removed.
         */
        public static void SetEscherProperty(EscherOptRecord opt, short propId, int value)
        {
            List<EscherProperty> props = opt.EscherProperties;
            for (List<EscherProperty>.Enumerator iterator = props.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherProperty prop = (EscherProperty)iterator.Current;
                if (prop.Id == propId)
                {
                    props.Remove(iterator.Current);
                }
            }
            if (value != -1)
            {
                opt.AddEscherProperty(new EscherSimpleProperty(propId, value));
                opt.SortProperties();
            }
        }

        /**
         * Set an simple escher property for this shape.
         *
         * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
         * @param value     value of the property. If value = -1 then the property is Removed.
         */
        public void SetEscherProperty(short propId, int value)
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            SetEscherProperty(opt, propId, value);
        }

        /**
         * Get the value of a simple escher property for this shape.
         *
         * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
         */
        public int GetEscherProperty(short propId)
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            EscherSimpleProperty prop = (EscherSimpleProperty)GetEscherProperty(opt, propId);
            return prop == null ? 0 : prop.PropertyValue;
        }

        /**
         * Get the value of a simple escher property for this shape.
         *
         * @param propId    The id of the property. One of the constants defined in EscherOptRecord.
         */
        public int GetEscherProperty(short propId, int defaultValue)
        {
            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(_escherContainer, EscherOptRecord.RECORD_ID);
            EscherSimpleProperty prop = (EscherSimpleProperty)GetEscherProperty(opt, propId);
            return prop == null ? defaultValue : prop.PropertyValue;
        }

        /**
         * @return  The shape Container and it's children that can represent this
         *          shape.
         */
        public EscherContainerRecord GetSpContainer()
        {
            return _escherContainer;
        }

        /**
         * Event which fires when a shape is inserted in the sheet.
         * In some cases we need to propagate Changes to upper level Containers.
         * <br>
         * Default implementation does nothing.
         *
         * @param sh - owning shape
         */
        protected void AfterInsert(Sheet sh)
        {

        }

        /**
         *  @return the <code>SlideShow</code> this shape belongs to
         */
        public Sheet Sheet
        {
            get
            {
                return _sheet;
            }
            set
            {
                _sheet = value;
            }
        }

        protected Color GetColor(int rgb, int alpha)
        {
            if (rgb >= 0x8000000)
            {
                int idx = rgb - 0x8000000;
                ColorSchemeAtom ca = Sheet.GetColorScheme();
                if (idx >= 0 && idx <= 7) rgb = ca.GetColor(idx);
            }
            Color tmp = Color.FromArgb(rgb);
            return Color.FromArgb(alpha, tmp.B, tmp.G, tmp.R);
        }

        /**
         * @return id for the shape.
         */
        public int GetShapeId()
        {
            EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            return spRecord == null ? 0 : spRecord.ShapeId;
        }

        /**
         * Sets shape ID
         *
         * @param id of the shape
         */
        public void SetShapeId(int id)
        {
            EscherSpRecord spRecord = _escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
            if (spRecord != null) spRecord.ShapeId = (id);
        }

        /**
         * Fill properties of this shape
         *
         * @return fill properties of this shape
         */
        public Fill GetFill()
        {
            if (_fill == null) _fill = new Fill(this);
            return _Fill;
        }


        /**
         * Returns the hyperlink assigned to this shape
         *
         * @return the hyperlink assigned to this shape
         * or <code>null</code> if not found.
         */
        public Hyperlink GetHyperlink()
        {
            return Hyperlink.Find(this);
        }

        public void Draw(Graphics graphics)
        {
            logger.Log(POILogger.INFO, "Rendering " + GetShapeName());
        }

        /**
         * Return shape outline as a java.awt.Shape object
         *
         * @return the shape outline
         */
        public Shape GetOutline()
        {
            return GetLogicalAnchor2D();
        }
    }


}


