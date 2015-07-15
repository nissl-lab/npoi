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
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using System;
    using NPOI.SS.UserModel;
    /// <summary>
    /// Represents a simple shape such as a line, rectangle or oval.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    [Serializable]
    public class HSSFSimpleShape: HSSFShape
    {
        // The commented out ones haven't been tested yet or aren't supported
        // by HSSFSimpleShape.

        public const short OBJECT_TYPE_LINE = (short)HSSFShapeTypes.Line;
        public const short OBJECT_TYPE_RECTANGLE = (short)HSSFShapeTypes.Rectangle;
        public const short OBJECT_TYPE_OVAL = (short)HSSFShapeTypes.Ellipse;
        public const short OBJECT_TYPE_ARC = (short)HSSFShapeTypes.Arc;
        //    public static short       OBJECT_TYPE_CHART              = 5;
        //    public static short       OBJECT_TYPE_TEXT               = 6;
        //    public static short       OBJECT_TYPE_BUTTON             = 7;
        public const short OBJECT_TYPE_PICTURE = (short)HSSFShapeTypes.PictureFrame;
        //    public static short       OBJECT_TYPE_POLYGON            = 9;
        //    public static short       OBJECT_TYPE_CHECKBOX           = 11;
        //    public static short       OBJECT_TYPE_OPTION_BUTTON      = 12;
        //    public static short       OBJECT_TYPE_EDIT_BOX           = 13;
        //    public static short       OBJECT_TYPE_LABEL              = 14;
        //    public static short       OBJECT_TYPE_DIALOG_BOX         = 15;
        //    public static short       OBJECT_TYPE_SPINNER            = 16;
        //    public static short       OBJECT_TYPE_SCROLL_BAR         = 17;
        //    public static short       OBJECT_TYPE_LIST_BOX           = 18;
        //    public static short       OBJECT_TYPE_GROUP_BOX          = 19;
        public const short OBJECT_TYPE_COMBO_BOX = (short)HSSFShapeTypes.HostControl;
        public const short OBJECT_TYPE_COMMENT = (short)HSSFShapeTypes.TextBox;
        public const short OBJECT_TYPE_MICROSOFT_OFFICE_DRAWING = 30;

        public const int WRAP_SQUARE = 0;
        public const int WRAP_BY_POINTS = 1;
        public const int WRAP_NONE = 2;

        private TextObjectRecord _textObjectRecord;

        public HSSFSimpleShape(EscherContainerRecord spContainer, ObjRecord objRecord, TextObjectRecord textObjectRecord)
            : base(spContainer, objRecord)
        {
            this._textObjectRecord = textObjectRecord;
        }
        public HSSFSimpleShape(EscherContainerRecord spContainer, ObjRecord objRecord)
            : base(spContainer, objRecord)
        {
            
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFSimpleShape"/> class.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <param name="anchor">The anchor.</param>
        public HSSFSimpleShape(HSSFShape parent, HSSFAnchor anchor)
            :base(parent, anchor)
        {
            _textObjectRecord = CreateTextObjRecord();
        }

        /// <summary>
        /// Gets the shape type.
        /// </summary>
        /// <value>One of the OBJECT_TYPE_* constants.</value>
        /// @see #OBJECT_TYPE_LINE
        /// @see #OBJECT_TYPE_OVAL
        /// @see #OBJECT_TYPE_RECTANGLE
        /// @see #OBJECT_TYPE_PICTURE
        /// @see #OBJECT_TYPE_COMMENT
        public virtual int ShapeType 
        {
            get
            {
                EscherSpRecord spRecord = (EscherSpRecord)GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);
                return spRecord.ShapeType;
            }
            set 
            {
                CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)GetObjRecord().SubRecords[0];
                cod.ObjectType = CommonObjectType.MicrosoftOfficeDrawing;
                EscherSpRecord spRecord = (EscherSpRecord)GetEscherContainer().GetChildById(EscherSpRecord.RECORD_ID);
                spRecord.ShapeType = ((short)value);
            }
        }
        public int WrapText
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.TEXT__WRAPTEXT);
                return null == property ? WRAP_SQUARE : property.PropertyValue;
            }
            set 
            {
                SetPropertyValue(new EscherSimpleProperty(EscherProperties.TEXT__WRAPTEXT, false, false, value));
            }
        }
        protected internal TextObjectRecord GetTextObjectRecord()
        {
            return _textObjectRecord;
        }
        protected virtual TextObjectRecord CreateTextObjRecord()
        {
            TextObjectRecord obj = new TextObjectRecord();
            obj.HorizontalTextAlignment = HorizontalTextAlignment.Center;
            obj.VerticalTextAlignment = VerticalTextAlignment.Center;
            obj.IsTextLocked = (true);
            obj.TextOrientation = TextOrientation.None;
            obj.Str = (new HSSFRichTextString(""));
            return obj;
        }
        /// <summary>
        /// Get or set the rich text string used by this object.
        /// </summary>
        public virtual IRichTextString String
        {
            get
            {
                return _textObjectRecord.Str;
            }
            set
            {
                //TODO add other shape types which can not contain text
                if (ShapeType == 0 || ShapeType == OBJECT_TYPE_LINE)
                {
                    throw new InvalidOperationException("Cannot set text for shape type: " + ShapeType);
                }
                HSSFRichTextString rtr = (HSSFRichTextString)value;
                // If font is not set we must set the default one
                if (rtr.NumFormattingRuns == 0) rtr.ApplyFont((short)0);
                TextObjectRecord txo = GetOrCreateTextObjRecord();
                txo.Str = (rtr);
                if (value.String != null)
                {
                    SetPropertyValue(new EscherSimpleProperty(EscherProperties.TEXT__TEXTID, value.String.GetHashCode()));
                }
            }
        }

        internal override HSSFShape CloneShape()
        {
            TextObjectRecord txo = null;
            EscherContainerRecord spContainer = new EscherContainerRecord();
            byte[] inSp = GetEscherContainer().Serialize();
            spContainer.FillFields(inSp, 0, new DefaultEscherRecordFactory());
            ObjRecord obj = (ObjRecord)GetObjRecord().CloneViaReserialise();
            if (GetTextObjectRecord() != null && this.String != null && null != this.String.String)
            {
                txo = (TextObjectRecord)GetTextObjectRecord().CloneViaReserialise();
            }
            return new HSSFSimpleShape(spContainer, obj, txo);
        }
        internal override void AfterInsert(HSSFPatriarch patriarch)
        {
            EscherAggregate agg = patriarch.GetBoundAggregate();
            agg.AssociateShapeToObjRecord(GetEscherContainer().GetChildById(EscherClientDataRecord.RECORD_ID), GetObjRecord());

            if (null != GetTextObjectRecord())
            {
                agg.AssociateShapeToObjRecord(GetEscherContainer().GetChildById(EscherTextboxRecord.RECORD_ID), GetTextObjectRecord());
            }
        }
        internal override void AfterRemove(HSSFPatriarch patriarch)
        {
            patriarch.GetBoundAggregate().RemoveShapeToObjRecord(GetEscherContainer().GetChildById(EscherClientDataRecord.RECORD_ID));
            if (null != GetEscherContainer().GetChildById(EscherTextboxRecord.RECORD_ID))
            {
                patriarch.GetBoundAggregate().RemoveShapeToObjRecord(GetEscherContainer().GetChildById(EscherTextboxRecord.RECORD_ID));
            }
        }
        protected override EscherContainerRecord CreateSpContainer()
        {
            EscherContainerRecord spContainer = new EscherContainerRecord();
            spContainer.RecordId=EscherContainerRecord.SP_CONTAINER;
            spContainer.Options = ((short)0x000F);

            EscherSpRecord sp = new EscherSpRecord();
            sp.RecordId = (EscherSpRecord.RECORD_ID);
            sp.Flags = (EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE);
            sp.Version = ((short)0x2);

            EscherClientDataRecord clientData = new EscherClientDataRecord();
            clientData.RecordId = (EscherClientDataRecord.RECORD_ID);
            clientData.Options = ((short)(0x0000));

            EscherOptRecord optRecord = new EscherOptRecord();
            optRecord.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEDASHING, LINESTYLE_SOLID));
            optRecord.SetEscherProperty(new EscherBoolProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x00080008));
            //        optRecord.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEWIDTH, LINEWIDTH_DEFAULT));
            optRecord.SetEscherProperty(new EscherRGBProperty(EscherProperties.FILL__FILLCOLOR, FILL__FILLCOLOR_DEFAULT));
            optRecord.SetEscherProperty(new EscherRGBProperty(EscherProperties.LINESTYLE__COLOR, LINESTYLE__COLOR_DEFAULT));
            optRecord.SetEscherProperty(new EscherBoolProperty(EscherProperties.FILL__NOFILLHITTEST, NO_FILLHITTEST_FALSE));
            optRecord.SetEscherProperty(new EscherBoolProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x00080008));

            optRecord.SetEscherProperty(new EscherShapePathProperty(EscherProperties.GEOMETRY__SHAPEPATH, EscherShapePathProperty.COMPLEX));
            optRecord.SetEscherProperty(new EscherBoolProperty(EscherProperties.GROUPSHAPE__PRINT, 0x080000));
            optRecord.RecordId = EscherOptRecord.RECORD_ID;

            EscherTextboxRecord escherTextbox = new EscherTextboxRecord();
            escherTextbox.RecordId = (EscherTextboxRecord.RECORD_ID);
            escherTextbox.Options = (short)0x0000;

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(optRecord);
            spContainer.AddChildRecord(this.Anchor.GetEscherAnchor());
            spContainer.AddChildRecord(clientData);
            spContainer.AddChildRecord(escherTextbox);
            return spContainer;
        }
        protected override ObjRecord CreateObjRecord()
        {
            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.IsLocked=true;
            c.IsPrintable = true;
            c.IsAutoFill=true;
            c.IsAutoline=true;
            EndSubRecord e = new EndSubRecord();

            obj.AddSubRecord(c);
            obj.AddSubRecord(e);
            return obj;
        }


        private TextObjectRecord GetOrCreateTextObjRecord()
        {
            if (GetTextObjectRecord() == null)
            {
                _textObjectRecord = CreateTextObjRecord();
            }
            EscherTextboxRecord escherTextbox = (EscherTextboxRecord)GetEscherContainer().GetChildById(EscherTextboxRecord.RECORD_ID);
            if (null == escherTextbox)
            {
                escherTextbox = new EscherTextboxRecord();
                escherTextbox.RecordId = (EscherTextboxRecord.RECORD_ID);
                escherTextbox.Options = ((short)0x0000);
                GetEscherContainer().AddChildRecord(escherTextbox);
                Patriarch.GetBoundAggregate().AssociateShapeToObjRecord(escherTextbox, _textObjectRecord);
            }
            return _textObjectRecord;
        }

        public bool FlipVertical { get; set; }

        public bool FlipHorizontal { get; set; }
    }
}