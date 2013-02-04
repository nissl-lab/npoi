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

namespace NPOI.HSSF.Model
{
    using System;
    using NPOI.DDF;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;
    using System.Collections.Generic;

    /// <summary>
    /// Represents a cell comment.
    /// This class Converts highlevel model data from HSSFComment
    /// to low-level records.
    /// @author Yegor Kozlov
    /// </summary>
    [Obsolete]
    public class CommentShape : TextboxShape
    {

        private NoteRecord note;

        /// <summary>
        /// Creates the low-level records for a comment.
        /// </summary>
        /// <param name="hssfShape">The highlevel shape.</param>
        /// <param name="shapeId">The shape id to use for this shape.</param>
        public CommentShape(HSSFComment hssfShape, int shapeId)
            : base(hssfShape, shapeId)
        {


            note = CreateNoteRecord(hssfShape, shapeId);

            ObjRecord obj = ObjRecord;
            List<SubRecord> records = obj.SubRecords;
            int cmoIdx = 0;
            for (int i = 0; i < records.Count; i++)
            {
                Object r = records[i];

                if (r is CommonObjectDataSubRecord)
                {
                    //modify autoFill attribute inherited from <c>TextObjectRecord</c>
                    CommonObjectDataSubRecord cmo = (CommonObjectDataSubRecord)r;
                    cmo.IsAutoFill=(false);
                    cmoIdx = i;
                }
            }
            //Add NoteStructure sub record
            //we don't know it's format, for now the record data Is empty
            NoteStructureSubRecord u = new NoteStructureSubRecord();
            obj.AddSubRecord(cmoIdx + 1, u);
        }

        /// <summary>
        /// Creates the low level NoteRecord
        /// which holds the comment attributes.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="shapeId">The shape id.</param>
        /// <returns></returns>
        private NoteRecord CreateNoteRecord(HSSFComment shape, int shapeId)
        {
            NoteRecord note = new NoteRecord();
            note.Column = shape.Column;
            note.Row = shape.Row;
            note.Flags = (shape.Visible ? NoteRecord.NOTE_VISIBLE : NoteRecord.NOTE_HIDDEN);
            note.ShapeId = shapeId;
            note.Author = (shape.Author == null ? "" : shape.Author);
            return note;
        }

        /// <summary>
        /// Sets standard escher options for a comment.
        /// This method is responsible for Setting default background,
        /// shading and other comment properties.
        /// </summary>
        /// <param name="shape">The highlevel shape.</param>
        /// <param name="opt">The escher records holding the proerties</param>
        /// <returns>The number of escher options added</returns>
        protected override int AddStandardOptions(HSSFShape shape, EscherOptRecord opt)
        {
            base.AddStandardOptions(shape, opt);

            //Remove Unnecessary properties inherited from TextboxShape
            for (int i = 0; i < opt.EscherProperties.Count; i++ )
            {
                EscherProperty prop = opt.EscherProperties[i];
                switch (prop.Id)
                {
                    case EscherProperties.TEXT__TEXTLEFT:
                    case EscherProperties.TEXT__TEXTRIGHT:
                    case EscherProperties.TEXT__TEXTTOP:
                    case EscherProperties.TEXT__TEXTBOTTOM:
                    case EscherProperties.GROUPSHAPE__PRINT:
                    case EscherProperties.FILL__FILLBACKCOLOR:
                    case EscherProperties.LINESTYLE__COLOR:
                        opt.EscherProperties.Remove(prop);
                        i--;
                        break;
                }
            }

            HSSFComment comment = (HSSFComment)shape;
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.GROUPSHAPE__PRINT, comment.Visible ? 0x000A0000 : 0x000A0002));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.SHADOWSTYLE__SHADOWOBSURED, 0x00030003));
            opt.AddEscherProperty(new EscherSimpleProperty(EscherProperties.SHADOWSTYLE__COLOR, 0x00000000));
            opt.SortProperties();
            return opt.EscherProperties.Count;   // # options Added
        }

        /// <summary>
        /// Gets the NoteRecord holding the comment attributes
        /// </summary>
        /// <value>The NoteRecord</value>
        public NoteRecord NoteRecord
        {
            get
            {
                return note;
            }
        }
        protected override int GetCmoObjectId(int shapeId)
        {
            return shapeId;
        }
    }
}