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
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using NPOI.SS.UserModel;


    /// <summary>
    /// A textbox Is a shape that may hold a rich text string.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    [Serializable]
    public class HSSFTextbox : HSSFSimpleShape
    {
        public const short OBJECT_TYPE_TEXT = 6;

        public HSSFTextbox(EscherContainerRecord spContainer, ObjRecord objRecord, TextObjectRecord textObjectRecord)
            : base(spContainer, objRecord, textObjectRecord)
        {

        }

        HSSFRichTextString str = new HSSFRichTextString("");

        /// <summary>
        /// Construct a new textbox with the given parent and anchor.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <param name="anchor">One of HSSFClientAnchor or HSSFChildAnchor</param>
        public HSSFTextbox(HSSFShape parent, HSSFAnchor anchor)
            : base(parent, anchor)
        {
            HorizontalAlignment = HorizontalTextAlignment.Left;
            VerticalAlignment = VerticalTextAlignment.Top;
            this.String = (new HSSFRichTextString(""));
        }


        protected override ObjRecord CreateObjRecord()
        {
            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.ObjectType = CommonObjectType.Text;
            c.IsLocked = (true);
            c.IsPrintable = (true);
            c.IsAutoFill = (true);
            c.IsAutoline = (true);
            EndSubRecord e = new EndSubRecord();
            obj.AddSubRecord(c);
            obj.AddSubRecord(e);
            return obj;
        }

        protected override EscherContainerRecord CreateSpContainer()
        {
            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherClientDataRecord clientData = new EscherClientDataRecord();
            EscherTextboxRecord escherTextbox = new EscherTextboxRecord();

            spContainer.RecordId = (EscherContainerRecord.SP_CONTAINER);
            spContainer.Options = ((short)0x000F);
            sp.RecordId = (EscherSpRecord.RECORD_ID);
            sp.Options = ((short)((EscherAggregate.ST_TEXTBOX << 4) | 0x2));

            sp.Flags = (EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE);
            opt.RecordId = (EscherOptRecord.RECORD_ID);
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTID, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__WRAPTEXT, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__ANCHORTEXT, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GROUPSHAPE__PRINT, 0x00080000));

            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTLEFT, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTRIGHT, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTTOP, 0));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.TEXT__TEXTBOTTOM, 0));

            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEDASHING, LINESTYLE_SOLID));
            opt.SetEscherProperty(new EscherBoolProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x00080008));
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEWIDTH, LINEWIDTH_DEFAULT));
            opt.SetEscherProperty(new EscherRGBProperty(EscherProperties.FILL__FILLCOLOR, FILL__FILLCOLOR_DEFAULT));
            opt.SetEscherProperty(new EscherRGBProperty(EscherProperties.LINESTYLE__COLOR, LINESTYLE__COLOR_DEFAULT));
            opt.SetEscherProperty(new EscherBoolProperty(EscherProperties.FILL__NOFILLHITTEST, NO_FILLHITTEST_FALSE));
            opt.SetEscherProperty(new EscherBoolProperty(EscherProperties.GROUPSHAPE__PRINT, 0x080000));

            EscherRecord anchor = Anchor.GetEscherAnchor();
            clientData.RecordId = (EscherClientDataRecord.RECORD_ID);
            clientData.Options = ((short)0x0000);
            escherTextbox.RecordId = (EscherTextboxRecord.RECORD_ID);
            escherTextbox.Options = ((short)0x0000);

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);
            spContainer.AddChildRecord(escherTextbox);

            return spContainer;
        }

        internal override void AfterInsert(HSSFPatriarch patriarch)
        {
            EscherAggregate agg = patriarch.GetBoundAggregate();
            agg.AssociateShapeToObjRecord(GetEscherContainer().GetChildById(EscherClientDataRecord.RECORD_ID), GetObjRecord());
            if (GetTextObjectRecord() != null)
            {
                agg.AssociateShapeToObjRecord(GetEscherContainer().GetChildById(EscherTextboxRecord.RECORD_ID), GetTextObjectRecord());
            }
        }

        internal override HSSFShape CloneShape()
        {
            TextObjectRecord txo = GetTextObjectRecord() == null ? null : (TextObjectRecord)GetTextObjectRecord().CloneViaReserialise();
            EscherContainerRecord spContainer = new EscherContainerRecord();
            byte[] inSp = GetEscherContainer().Serialize();
            spContainer.FillFields(inSp, 0, new DefaultEscherRecordFactory());
            ObjRecord obj = (ObjRecord)GetObjRecord().CloneViaReserialise();
            return new HSSFTextbox(spContainer, obj, txo);
        }

        internal override void AfterRemove(HSSFPatriarch patriarch)
        {
            patriarch.GetBoundAggregate().RemoveShapeToObjRecord(GetEscherContainer().GetChildById(EscherClientDataRecord.RECORD_ID));
            patriarch.GetBoundAggregate().RemoveShapeToObjRecord(GetEscherContainer().GetChildById(EscherTextboxRecord.RECORD_ID));
        }


        /// <summary>
        /// Gets or sets the left margin within the textbox.
        /// </summary>
        /// <value>The margin left.</value>
        public int MarginLeft
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.TEXT__TEXTLEFT);
                return property == null ? 0 : property.PropertyValue;
            }
            set { SetPropertyValue(new EscherSimpleProperty(EscherProperties.TEXT__TEXTLEFT, value)); }
        }


        /// <summary>
        /// Gets or sets the right margin within the textbox.
        /// </summary>
        /// <value>The margin right.</value>
        public int MarginRight
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.TEXT__TEXTRIGHT);
                return property == null ? 0 : property.PropertyValue;
            }
            set { SetPropertyValue(new EscherSimpleProperty(EscherProperties.TEXT__TEXTRIGHT, value)); }
        }

        /// <summary>
        /// Gets or sets the top margin within the textbox
        /// </summary>
        /// <value>The top margin.</value>
        public int MarginTop
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.TEXT__TEXTTOP);
                return property == null ? 0 : property.PropertyValue;
            }
            set { SetPropertyValue(new EscherSimpleProperty(EscherProperties.TEXT__TEXTTOP, value)); }
        }

        /// <summary>
        /// Gets or sets the bottom margin within the textbox.
        /// </summary>
        /// <value>The margin bottom.</value>
        public int MarginBottom
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.TEXT__TEXTBOTTOM);
                return property == null ? 0 : property.PropertyValue;
            }
            set { SetPropertyValue(new EscherSimpleProperty(EscherProperties.TEXT__TEXTBOTTOM, value)); }
        }

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        /// <value>The horizontal alignment.</value>
        public HorizontalTextAlignment HorizontalAlignment
        {
            get { return GetTextObjectRecord().HorizontalTextAlignment; }
            set { GetTextObjectRecord().HorizontalTextAlignment = value; }
        }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        /// <value>The vertical alignment.</value>
        public VerticalTextAlignment VerticalAlignment
        {
            get { return GetTextObjectRecord().VerticalTextAlignment; }
            set { GetTextObjectRecord().VerticalTextAlignment = value; }
        }


        public override int ShapeType
        {
            get { return base.ShapeType; }
            set
            {
                throw new InvalidOperationException("Shape type can not be changed in " + this.GetType().Name);
            }
        }

    }
}
