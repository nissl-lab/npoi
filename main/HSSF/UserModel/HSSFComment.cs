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
    /// Represents a cell comment - a sticky note associated with a cell.
    /// @author Yegor Kozlov
    /// </summary>
    [Serializable]
    public class HSSFComment : HSSFTextbox, IComment
    {
        private const int FILL_TYPE_SOLID = 0;
        private const int FILL_TYPE_PICTURE = 3;

        private const int GROUP_SHAPE_PROPERTY_DEFAULT_VALUE = 655362;
        private const int GROUP_SHAPE_HIDDEN_MASK = 0x1000002;
        private const int GROUP_SHAPE_NOT_HIDDEN_MASK = unchecked((int)0xFEFFFFFD);

        private NoteRecord _note = null;

        //private TextObjectRecord txo = null;
        public HSSFComment(EscherContainerRecord spContainer, ObjRecord objRecord, TextObjectRecord textObjectRecord, NoteRecord _note)
            : base(spContainer, objRecord, textObjectRecord)
        {

            this._note = _note;
        }
        /// <summary>
        /// Construct a new comment with the given parent and anchor.
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="anchor">defines position of this anchor in the sheet</param>
        public HSSFComment(HSSFShape parent, HSSFAnchor anchor)
            : base(parent, anchor)
        {
            _note = CreateNoteRecord();

            //default color for comments
            this.FillColor = 0x08000050;

            //by default comments are hidden
            Visible = false;

            Author = "";
            CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)GetObjRecord().SubRecords[0];
            cod.ObjectType = CommonObjectType.Comment; 
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFComment"/> class.
        /// </summary>
        /// <param name="note">The note.</param>
        /// <param name="txo">The txo.</param>
        public HSSFComment(NoteRecord note, TextObjectRecord txo)
            : this((HSSFShape)null, new HSSFClientAnchor())
        {
            this._note = note;
        }


        internal override void AfterInsert(HSSFPatriarch patriarch)
        {
            base.AfterInsert(patriarch);
            patriarch.GetBoundAggregate().AddTailRecord(NoteRecord);
        }

        protected override EscherContainerRecord CreateSpContainer()
        {
            EscherContainerRecord spContainer = base.CreateSpContainer();
            EscherOptRecord opt = (EscherOptRecord)spContainer.GetChildById(EscherOptRecord.RECORD_ID);
            opt.RemoveEscherProperty(EscherProperties.TEXT__TEXTLEFT);
            opt.RemoveEscherProperty(EscherProperties.TEXT__TEXTRIGHT);
            opt.RemoveEscherProperty(EscherProperties.TEXT__TEXTTOP);
            opt.RemoveEscherProperty(EscherProperties.TEXT__TEXTBOTTOM);
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.GROUPSHAPE__PRINT, false, false, GROUP_SHAPE_PROPERTY_DEFAULT_VALUE));
            return spContainer;
        }


        protected override ObjRecord CreateObjRecord()
        {
            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.ObjectType = (CommonObjectType)OBJECT_TYPE_COMMENT;
            c.IsLocked = (true);
            c.IsPrintable = (true);
            c.IsAutoFill = (false);
            c.IsAutoline = (true);

            NoteStructureSubRecord u = new NoteStructureSubRecord();
            EndSubRecord e = new EndSubRecord();
            obj.AddSubRecord(c);
            obj.AddSubRecord(u);
            obj.AddSubRecord(e);
            return obj;
        }

        private NoteRecord CreateNoteRecord()
        {
            NoteRecord note = new NoteRecord();
            note.Flags = (NoteRecord.NOTE_HIDDEN);
            note.Author = ("");
            return note;
        }

        public override int ShapeId
        {
            get { return base.ShapeId; }
            set
            {
                if (value > 65535)
                    throw new ArgumentException("Cannot add more than 65535 shapes");
                base.ShapeId = (value);
                CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)GetObjRecord().SubRecords[0];
                cod.ObjectId = value;
                _note.ShapeId = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="HSSFComment"/> is visible.
        /// </summary>
        /// <value><c>true</c> if visible; otherwise, <c>false</c>.</value>
        /// Sets whether this comment Is visible.
        /// @return 
        /// <c>true</c>
        ///  if the comment Is visible, 
        /// <c>false</c>
        ///  otherwise
        public bool Visible
        {
            get
            {
                return _note.Flags == NoteRecord.NOTE_VISIBLE;
            }
            set
            {
                if (_note != null) _note.Flags = value ? NoteRecord.NOTE_VISIBLE : NoteRecord.NOTE_HIDDEN;
                SetHidden(!value);
            }
        }

        /// <summary>
        /// Gets or sets the row of the cell that Contains the comment
        /// </summary>
        /// <value>the 0-based row of the cell that Contains the comment</value>
        public int Row
        {
            get { return _note.Row; }
            set
            {
                if (_note != null) _note.Row = value;
            }
        }


        /// <summary>
        /// Gets or sets the column of the cell that Contains the comment
        /// </summary>
        /// <value>the 0-based column of the cell that Contains the comment</value>
        public int Column
        {
            get { return _note.Column; }
            set
            {
                if (_note != null) _note.Column = value;
            }
        }

        /// <summary>
        /// Gets or sets the name of the original comment author
        /// </summary>
        /// <value>the name of the original author of the comment</value>
        public String Author
        {
            get
            {
                return _note.Author;
            }
            set
            {
                if (_note != null) _note.Author = value;
            }
        }

        

        /// <summary>
        /// Gets the note record.
        /// </summary>
        /// <value>the underlying Note record.</value>
        internal NoteRecord NoteRecord
        {
            get { return _note; }
        }

        /**
         * Do we know which cell this comment belongs to?
         */
        public bool HasPosition
        {
            get
            {
                if (_note == null) return false;
                if (this.Column < 0 || this.Row < 0) return false;
                return true;
            }
        }

        public IClientAnchor ClientAnchor
        {
            get
            {
                HSSFAnchor ha = base.Anchor;
                if (ha is IClientAnchor)
                {
                    return (IClientAnchor)ha;
                }

                throw new InvalidCastException("Anchor can not be changed in "
                        + typeof(IClientAnchor).Name);
            }
        }

        public override int ShapeType
        {
            get
            {
                return base.ShapeType;
            }
            set
            {
                throw new InvalidOperationException("Shape type can not be changed in " + this.GetType().Name);
            }
        }
        

        internal override void AfterRemove(HSSFPatriarch patriarch)
        {
            base.AfterRemove(patriarch);
            patriarch.GetBoundAggregate().RemoveTailRecord(this.NoteRecord);
        }
        internal override HSSFShape CloneShape()
        {
            TextObjectRecord txo = (TextObjectRecord)GetTextObjectRecord().CloneViaReserialise();
            EscherContainerRecord spContainer = new EscherContainerRecord();
            byte[] inSp = GetEscherContainer().Serialize();
            spContainer.FillFields(inSp, 0, new DefaultEscherRecordFactory());
            ObjRecord obj = (ObjRecord)GetObjRecord().CloneViaReserialise();
            NoteRecord note = (NoteRecord)NoteRecord.CloneViaReserialise();
            return new HSSFComment(spContainer, obj, txo, note);
        }
        public void SetBackgroundImage(int pictureIndex)
        {
            SetPropertyValue(new EscherSimpleProperty(EscherProperties.FILL__PATTERNTEXTURE, false, true, pictureIndex));
            SetPropertyValue(new EscherSimpleProperty(EscherProperties.FILL__FILLTYPE, false, false, FILL_TYPE_PICTURE));
            EscherBSERecord bse = ((HSSFWorkbook)((HSSFPatriarch)Patriarch).Sheet.Workbook).Workbook.GetBSERecord(pictureIndex);
            bse.Ref = (bse.Ref + 1);
        }

        public void ResetBackgroundImage()
        {
            EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.FILL__PATTERNTEXTURE);
            if (null != property)
            {
                EscherBSERecord bse = ((HSSFWorkbook)((HSSFPatriarch)Patriarch).Sheet.Workbook).Workbook.GetBSERecord(property.PropertyValue);
                bse.Ref = (bse.Ref - 1);
                GetOptRecord().RemoveEscherProperty(EscherProperties.FILL__PATTERNTEXTURE);
            }
            SetPropertyValue(new EscherSimpleProperty(EscherProperties.FILL__FILLTYPE, false, false, FILL_TYPE_SOLID));
        }

        public int GetBackgroundImageId()
        {
            EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.FILL__PATTERNTEXTURE);
            return property == null ? 0 : property.PropertyValue;
        }
        private void SetHidden(bool value)
        {
            EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.GROUPSHAPE__PRINT);
            // see http://msdn.microsoft.com/en-us/library/dd949807(v=office.12).aspx
            if (value)
            {
                SetPropertyValue(new EscherSimpleProperty(EscherProperties.GROUPSHAPE__PRINT, false, false, property.PropertyValue | GROUP_SHAPE_HIDDEN_MASK));
            }
            else
            {
                SetPropertyValue(new EscherSimpleProperty(EscherProperties.GROUPSHAPE__PRINT, false, false, property.PropertyValue & GROUP_SHAPE_NOT_HIDDEN_MASK));
            }
        }
    }
}