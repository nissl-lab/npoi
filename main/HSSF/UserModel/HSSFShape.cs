/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;
    using NPOI.SS.UserModel;
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using System.IO;
    using NPOI.Util;
    /// <summary>
    /// An abstract shape.
    /// 
    /// Note: Microsoft Excel seems to sometimes disallow 
    /// higher y1 than y2 or higher x1 than x2 in the anchor, you might need to 
    /// reverse them and draw shapes vertically or horizontally flipped! 
    /// </summary>
    [Serializable]
    public abstract class HSSFShape //: IShape
    {
        public const int LINEWIDTH_ONE_PT = 12700; // 12700 = 1pt
        public const int LINEWIDTH_DEFAULT = 9525;
        public const int LINESTYLE__COLOR_DEFAULT = 0x08000040;
        public const int FILL__FILLCOLOR_DEFAULT = 0x08000009;
        public const bool NO_FILL_DEFAULT = true;

        public const int LINESTYLE_SOLID = 0;              // Solid (continuous) pen
        public const int LINESTYLE_DASHSYS = 1;            // PS_DASH system   dash style
        public const int LINESTYLE_DOTSYS = 2;             // PS_DOT system   dash style
        public const int LINESTYLE_DASHDOTSYS = 3;         // PS_DASHDOT system dash style
        public const int LINESTYLE_DASHDOTDOTSYS = 4;      // PS_DASHDOTDOT system dash style
        public const int LINESTYLE_DOTGEL = 5;             // square dot style
        public const int LINESTYLE_DASHGEL = 6;            // dash style
        public const int LINESTYLE_LONGDASHGEL = 7;        // long dash style
        public const int LINESTYLE_DASHDOTGEL = 8;         // dash short dash
        public const int LINESTYLE_LONGDASHDOTGEL = 9;     // long dash short dash
        public const int LINESTYLE_LONGDASHDOTDOTGEL = 10; // long dash short dash short dash
        public const int LINESTYLE_NONE = -1;

        public const int LINESTYLE_DEFAULT = LINESTYLE_NONE;

        public const int NO_FILLHITTEST_TRUE = 0x00110000;
        public const int NO_FILLHITTEST_FALSE = 0x00010000;

        HSSFShape parent;
        [NonSerialized]
        protected HSSFAnchor anchor;
        [NonSerialized]
        protected internal HSSFPatriarch _patriarch;

        private EscherContainerRecord _escherContainer;
        private ObjRecord _objRecord;
        private EscherOptRecord _optRecord;
        //int lineStyleColor = 0x08000040;
        //int fillColor = 0x08000009;
        //int lineWidth = LINEWIDTH_DEFAULT;
        //LineStyle lineStyle = LineStyle.Solid;
        //bool noFill = false;

        /**
         * creates shapes from existing file
         * @param spContainer
         * @param objRecord
         */
        public HSSFShape(EscherContainerRecord spContainer, ObjRecord objRecord)
        {
            this._escherContainer = spContainer;
            this._objRecord = objRecord;
            this._optRecord = (EscherOptRecord)spContainer.GetChildById(EscherOptRecord.RECORD_ID);
            this.anchor = HSSFAnchor.CreateAnchorFromEscher(spContainer);
        }

        /// <summary>
        /// Create a new shape with the specified parent and anchor.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <param name="anchor">The anchor.</param>
        public HSSFShape(HSSFShape parent, HSSFAnchor anchor)
        {
            this.parent = parent;
            this.anchor = anchor;
            this._escherContainer = CreateSpContainer();
            _optRecord = (EscherOptRecord)_escherContainer.GetChildById(EscherOptRecord.RECORD_ID);
            _objRecord = CreateObjRecord();
        }
        protected abstract EscherContainerRecord CreateSpContainer();

        protected abstract ObjRecord CreateObjRecord();
        internal abstract void AfterRemove(HSSFPatriarch patriarch);
        internal abstract void AfterInsert(HSSFPatriarch patriarch);

        public virtual int ShapeId
        {
            get
            {
                return ((EscherSpRecord)_escherContainer.GetChildById(EscherSpRecord.RECORD_ID)).ShapeId;
            }
            set
            {
                EscherSpRecord spRecord = (EscherSpRecord)_escherContainer.GetChildById(EscherSpRecord.RECORD_ID);
                spRecord.ShapeId = value;
                CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)_objRecord.SubRecords[0];
                cod.ObjectId = (short)(value % 1024);
            }
        }
        /// <summary>
        /// Gets the parent shape.
        /// </summary>
        /// <value>The parent.</value>
        public HSSFShape Parent
        {
            get { return this.parent; }
            set { this.parent = value; }
        }

        /// <summary>
        /// Gets or sets the anchor that is used by this shape.
        /// </summary>
        /// <value>The anchor.</value>
        public HSSFAnchor Anchor
        {
            get { return anchor; }
            set
            {
                int i = 0;
                int recordId = -1;
                if (parent == null)
                {
                    if (value is HSSFChildAnchor)
                        throw new ArgumentException("Must use client anchors for shapes directly attached to sheet.");
                    EscherClientAnchorRecord anch = (EscherClientAnchorRecord)_escherContainer.GetChildById(EscherClientAnchorRecord.RECORD_ID);
                    if (null != anch)
                    {
                        for (i = 0; i < _escherContainer.ChildRecords.Count; i++)
                        {
                            if (_escherContainer.GetChild(i).RecordId == EscherClientAnchorRecord.RECORD_ID)
                            {
                                if (i != _escherContainer.ChildRecords.Count - 1)
                                {
                                    recordId = _escherContainer.GetChild(i + 1).RecordId;
                                }
                            }
                        }
                        _escherContainer.RemoveChildRecord(anch);
                    }
                }
                else
                {
                    if (value is HSSFClientAnchor)
                        throw new ArgumentException("Must use child anchors for shapes attached to Groups.");
                    EscherChildAnchorRecord anch = (EscherChildAnchorRecord)_escherContainer.GetChildById(EscherChildAnchorRecord.RECORD_ID);
                    if (null != anch)
                    {
                        for (i = 0; i < _escherContainer.ChildRecords.Count; i++)
                        {
                            if (_escherContainer.GetChild(i).RecordId == EscherChildAnchorRecord.RECORD_ID)
                            {
                                if (i != _escherContainer.ChildRecords.Count - 1)
                                {
                                    recordId = _escherContainer.GetChild(i + 1).RecordId;
                                }
                            }
                        }
                        _escherContainer.RemoveChildRecord(anch);
                    }
                }
                if (-1 == recordId)
                {
                    _escherContainer.AddChildRecord(value.GetEscherAnchor());
                }
                else
                {
                    _escherContainer.AddChildBefore(value.GetEscherAnchor(), recordId);
                }
                this.anchor = value;
            }
        }

        //public void SetLineStyleColor(int lineStyleColor)
        //{
        //    this.lineStyleColor = lineStyleColor;
        //}
        /// <summary>
        /// The color applied to the lines of this shape.
        /// </summary>
        /// <value>The color of the line style.</value>
        public int LineStyleColor
        {
            get
            {
                //return lineStyleColor;
                EscherRGBProperty rgbProperty = (EscherRGBProperty)_optRecord.Lookup(EscherProperties.LINESTYLE__COLOR);
                return rgbProperty == null ? LINESTYLE__COLOR_DEFAULT : rgbProperty.RgbColor;
            }
            set
            {
                SetPropertyValue(new EscherRGBProperty(EscherProperties.LINESTYLE__COLOR, value));
            }
        }
        internal EscherContainerRecord GetEscherContainer()
        {
            return _escherContainer;
        }
        /// <summary>
        /// Sets the color applied to the lines of this shape
        /// </summary>
        /// <param name="red">The red.</param>
        /// <param name="green">The green.</param>
        /// <param name="blue">The blue.</param>
        public void SetLineStyleColor(int red, int green, int blue)
        {
            int lineStyleColor = ((blue) << 16) | ((green) << 8) | red;
            SetPropertyValue(new EscherRGBProperty(EscherProperties.LINESTYLE__COLOR, lineStyleColor));
        }
        protected void SetPropertyValue(EscherProperty property)
        {
            _optRecord.SetEscherProperty(property);
        }
        /// <summary>
        /// Gets or sets the color used to fill this shape.
        /// </summary>
        /// <value>The color of the fill.</value>
        public int FillColor
        {
            get 
            {
                EscherRGBProperty rgbProperty = (EscherRGBProperty)_optRecord.Lookup(EscherProperties.FILL__FILLCOLOR);
                return rgbProperty == null ? FILL__FILLCOLOR_DEFAULT : rgbProperty.RgbColor;
            }
            set 
            {
                SetPropertyValue(new EscherRGBProperty(EscherProperties.FILL__FILLCOLOR, value));
            }
        }

        /// <summary>
        /// Sets the color used to fill this shape.
        /// </summary>
        /// <param name="red">The red.</param>
        /// <param name="green">The green.</param>
        /// <param name="blue">The blue.</param>
        public void SetFillColor(int red, int green, int blue)
        {
            int fillColor = ((blue) << 16) | ((green) << 8) | red;
            SetPropertyValue(new EscherRGBProperty(EscherProperties.FILL__FILLCOLOR, fillColor));
        }

        /// <summary>
        /// Gets or sets with width of the line in EMUs.  12700 = 1 pt.
        /// </summary>
        /// <value>The width of the line.</value>
        public int LineWidth
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)_optRecord.Lookup(EscherProperties.LINESTYLE__LINEWIDTH);
                return property == null ? LINEWIDTH_DEFAULT : property.PropertyValue;
            }
            set
            {
                SetPropertyValue(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEWIDTH, value));
            }
        }

        /// <summary>
        /// Gets or sets One of the constants in LINESTYLE_*
        /// </summary>
        /// <value>The line style.</value>
        public LineStyle LineStyle
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)_optRecord.Lookup(EscherProperties.LINESTYLE__LINEDASHING);
                if (null == property)
                {
                    return (LineStyle)LINESTYLE_DEFAULT;
                }
                return (LineStyle)property.PropertyValue;
            }
            set
            {
                SetPropertyValue(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEDASHING, (int)value));
                if ((int)LineStyle != HSSFShape.LINESTYLE_SOLID)
                {
                    SetPropertyValue(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEENDCAPSTYLE, 0));
                    if ((int)LineStyle == HSSFShape.LINESTYLE_NONE)
                    {
                        SetPropertyValue(new EscherBoolProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x00080000));
                    }
                    else
                    {
                        SetPropertyValue(new EscherBoolProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x00080008));
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is no fill.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this shape Is not filled with a color; otherwise, <c>false</c>.
        /// </value>
        public bool IsNoFill
        {
            get
            {
                EscherBoolProperty property = (EscherBoolProperty)_optRecord.Lookup(EscherProperties.FILL__NOFILLHITTEST);
                return property == null ? NO_FILL_DEFAULT : property.PropertyValue == NO_FILLHITTEST_TRUE;
            }
            set
            {
                SetPropertyValue(new EscherBoolProperty(EscherProperties.FILL__NOFILLHITTEST, value ? NO_FILLHITTEST_TRUE : NO_FILLHITTEST_FALSE));
            }
        }

        /// <summary>
        /// whether this shape is vertically flipped.
        /// </summary>
        public bool IsFlipVertical
        {
            get
            {
                EscherSpRecord sp = (EscherSpRecord)GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);
                return (sp.Flags & EscherSpRecord.FLAG_FLIPVERT) != 0;
            }
            set
            {
                EscherSpRecord sp = (EscherSpRecord)GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);
                if (value)
                {
                    sp.Flags = (sp.Flags | EscherSpRecord.FLAG_FLIPVERT);
                }
                else
                {
                    sp.Flags = (sp.Flags & (int.MaxValue - EscherSpRecord.FLAG_FLIPVERT));
                }
            }
        }

        /// <summary>
        /// whether this shape is horizontally flipped.
        /// </summary>
        public bool IsFlipHorizontal
        {
            get
            {
                EscherSpRecord sp = (EscherSpRecord)GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);
                return (sp.Flags & EscherSpRecord.FLAG_FLIPHORIZ) != 0;
            }
            set
            {
                EscherSpRecord sp = (EscherSpRecord)GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);
                if (value)
                {
                    sp.Flags=(sp.Flags | EscherSpRecord.FLAG_FLIPHORIZ);
                }
                else
                {
                    sp.Flags=(sp.Flags & (int.MaxValue - EscherSpRecord.FLAG_FLIPHORIZ));
                }
            }
        }

        /// <summary>
        /// get or set the rotation, in degrees, that is applied to a shape.
        /// Negative values specify rotation in the counterclockwise direction.
        /// Rotation occurs around the center of the shape.
        /// The default value for this property is 0x00000000
        /// </summary>
        public int RotationDegree
        {
            get
            {
                using (MemoryStream bos = new MemoryStream())
                {
                    EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.TRANSFORM__ROTATION);
                    if (null == property)
                    {
                        return 0;
                    }
                    try
                    {
                        LittleEndian.PutInt(property.PropertyValue, bos);
                        return LittleEndian.GetShort(bos.ToArray(), 2);
                    }
                    catch (IOException)
                    {
                        //e.printStackTrace();
                        return 0;
                    }
                }
            }
            set
            {
                SetPropertyValue(new EscherSimpleProperty(EscherProperties.TRANSFORM__ROTATION, (value << 16)));
            }
        }

        /// <summary>
        /// Count of all children and their childrens children.
        /// </summary>
        /// <value>The count of all children.</value>
        public virtual int CountOfAllChildren
        {
            get { return 1; }
        }

        internal abstract HSSFShape CloneShape();

        public HSSFPatriarch Patriarch
        {
            get
            {
                return _patriarch;
            }
            set
            {
                this._patriarch = value;
            }
        }

        protected internal ObjRecord GetObjRecord()
        {
            return _objRecord;
        }

        protected internal EscherOptRecord GetOptRecord()
        {
            return _optRecord;
        }
    }
}