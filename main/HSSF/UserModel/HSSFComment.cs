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
    using System.Collections;

    using NPOI.Util;
    using NPOI.DDF;
    using NPOI.HSSF.Record;
    using NPOI.SS.UserModel;


    /// <summary>
    /// Represents a cell comment - a sticky note associated with a cell.
    /// @author Yegor Kozlov
    /// </summary>
    [Serializable]
    public class HSSFComment : HSSFTextbox,IComment
    {

        private bool visible;
        private int col, row;
        private String author;

        
        private NoteRecord note = null;
        
        private TextObjectRecord txo = null;

        /// <summary>
        /// Construct a new comment with the given parent and anchor.
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="anchor">defines position of this anchor in the sheet</param>
        public HSSFComment(HSSFShape parent, HSSFAnchor anchor):base(parent, anchor)
        {
            
            this.ShapeType = (OBJECT_TYPE_COMMENT);

            //default color for comments
            this.FillColor = 0x08000050;

            //by default comments are hidden
            visible = false;

            author = "";
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFComment"/> class.
        /// </summary>
        /// <param name="note">The note.</param>
        /// <param name="txo">The txo.</param>
        public HSSFComment(NoteRecord note, TextObjectRecord txo):this((HSSFShape)null, (HSSFAnchor)null)
        {          
            this.txo = txo;
            this.note = note;
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
                return this.visible;
            }
            set {
                if (note != null) note.Flags = value ? NoteRecord.NOTE_VISIBLE : NoteRecord.NOTE_HIDDEN;
                this.visible = value;
            }
        }

        /// <summary>
        /// Gets or sets the row of the cell that Contains the comment
        /// </summary>
        /// <value>the 0-based row of the cell that Contains the comment</value>
        public int Row
        {
            get{return row;}
            set{            
                if (note != null) note.Row=value;
            this.row = value;}
        }


        /// <summary>
        /// Gets or sets the column of the cell that Contains the comment
        /// </summary>
        /// <value>the 0-based column of the cell that Contains the comment</value>
        public int Column
        {
            get { return col; }
            set
            {
                if (note != null) note.Column=value;
                this.col = value;
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
                return author;
            }
            set
            {
                if (note != null) note.Author = value;
                this.author = value;
            }
        }

        /// <summary>
        /// Gets or sets the rich text string used by this comment.
        /// </summary>   
        public override IRichTextString String
        {
            get { return base.String; }
            set
            {
                //if font Is not Set we must Set the default one
                if (value.NumFormattingRuns == 0) value.ApplyFont((short)0);

                if (txo != null)
                {
                    txo.Str=value;
                }
                base.String = value;
            }
        }

        /// <summary>
        /// Gets the note record.
        /// </summary>
        /// <value>the underlying Note record.</value>
        internal NoteRecord NoteRecord
        {
            get { return note; }
        }
        /// <summary>
        /// Gets the text object record.
        /// </summary>
        /// <value>the underlying Text record</value>
        internal TextObjectRecord TextObjectRecord
        {
            get { return txo; }
        }

    }
}