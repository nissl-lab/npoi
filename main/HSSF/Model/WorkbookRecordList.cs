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
    using System.Collections;
    using NPOI.HSSF.Record;
    using System.Collections.Generic;


    /// <summary>
    /// List for records in Workbook
    /// </summary>
    public class WorkbookRecordList
    {
        private List<Record> records = new List<Record>();

        private int protpos = 0;   // holds the position of the protect record.
        private int bspos = 0;   // holds the position of the last bound sheet.
        private int tabpos = 0;   // holds the position of the tabid record
        private int fontpos = 0;   // hold the position of the last font record
        private int xfpos = 0;   // hold the position of the last extended font record
        private int backuppos = 0;   // holds the position of the backup record.
        private int namepos = 0;   // holds the position of last name record
        private int supbookpos = 0;   // holds the position of sup book
        private int externsheetPos = 0;// holds the position of the extern sheet
        private int palettepos = -1;   // hold the position of the palette, if applicable


        /// <summary>
        /// Gets or sets the records.
        /// </summary>
        /// <value>The records.</value>
        public List<Record> Records
        {
            get { return records; }
            set { this.records = value; }
        }

        /// <summary>
        /// Gets the count.
        /// </summary>
        /// <value>The count.</value>
        public int Count
        {
            get { return records.Count; }
        }

        /// <summary>
        /// Gets the <see cref="NPOI.HSSF.Record.Record"/> at the specified index.
        /// </summary>
        /// <value></value>
        public Record this[int index]
        {
            get { return (Record)records[index]; }

        }

        /// <summary>
        /// Adds the specified pos.
        /// </summary>
        /// <param name="pos">The pos.</param>
        /// <param name="r">The r.</param>
        public void Add(int pos, Record r)
        {
            records.Insert(pos, r);
            if (Protpos >= pos) Protpos=(protpos + 1);
            if (Bspos >= pos) Bspos=(bspos + 1);
            if (Tabpos >= pos) Tabpos=(tabpos + 1);
            if (Fontpos >= pos) Fontpos=(fontpos + 1);
            if (Xfpos >= pos) Xfpos=(xfpos + 1);
            if (Backuppos >= pos) Backuppos=(backuppos + 1);
            if (Namepos>= pos) Namepos=(namepos + 1);
            if (Supbookpos >= pos) Supbookpos=(supbookpos + 1);
            if ((Palettepos!= -1) && (Palettepos>= pos)) Palettepos=(palettepos + 1);
            if (ExternsheetPos >= pos) ExternsheetPos=ExternsheetPos + 1;
        }
        public IEnumerator<Record> GetEnumerator()
        {
            return records.GetEnumerator();
        }

        /// <summary>
        /// Removes the specified record.
        /// </summary>
        /// <param name="record">The record.</param>
        public void Remove(Record record)
        {
            int i = records.IndexOf(record);
            this.Remove(i);
        }

        /// <summary>
        /// Removes the specified position.
        /// </summary>
        /// <param name="pos">The position.</param>
        public void Remove(int pos)
        {
            records.RemoveAt(pos);
            if (Protpos >= pos) Protpos=protpos - 1;
            if (Bspos >= pos) Bspos=bspos - 1;
            if (Tabpos >= pos) Tabpos=tabpos - 1;
            if (Fontpos >= pos) Fontpos=fontpos - 1;
            if (Xfpos >= pos) Xfpos=xfpos - 1;
            if (Backuppos >= pos) Backuppos=backuppos - 1;
            if (Namepos >= pos) Namepos=Namepos - 1;
            if (Supbookpos >= pos) Supbookpos=Supbookpos - 1;
            if ((Palettepos != -1) && (Palettepos >= pos)) Palettepos=palettepos - 1;
            if (ExternsheetPos >= pos) ExternsheetPos=ExternsheetPos- 1;
        }

        /// <summary>
        /// Gets or sets the protpos.
        /// </summary>
        /// <value>The protpos.</value>
        public int Protpos
        {
            get { return protpos; }
            set { this.protpos = value; }
        }

        /// <summary>
        /// Gets or sets the bspos.
        /// </summary>
        /// <value>The bspos.</value>
        public int Bspos
        {
            get{return bspos;}
            set { this.bspos = value; }
        }

        /// <summary>
        /// Gets or sets the tabpos.
        /// </summary>
        /// <value>The tabpos.</value>
        public int Tabpos
        {
            get { return tabpos; }
            set { this.tabpos =value; }
        }

        /// <summary>
        /// Gets or sets the fontpos.
        /// </summary>
        /// <value>The fontpos.</value>
        public int Fontpos
        {
            get { return fontpos; }
            set { this.fontpos = value; }
        }
        /// <summary>
        /// Gets or sets the xfpos.
        /// </summary>
        /// <value>The xfpos.</value>
        public int Xfpos
        {
            get { return xfpos; }
            set { this.xfpos = value; }
        }

        /// <summary>
        /// Gets or sets the backuppos.
        /// </summary>
        /// <value>The backuppos.</value>
        public int Backuppos
        {
            get{return backuppos;}
            set { this.backuppos = value; }
        }

        /// <summary>
        /// Gets or sets the palettepos.
        /// </summary>
        /// <value>The palettepos.</value>
        public int Palettepos
        {
            get{return palettepos;}
            set { this.palettepos = value; }
        }


        /// <summary>
        /// Gets or sets the namepos.
        /// </summary>
        /// <value>The namepos.</value>
        public int Namepos
        {
            get { return namepos; }
            set { this.namepos = value; }
        }

        /// <summary>
        /// Gets or sets the supbookpos.
        /// </summary>
        /// <value>The supbookpos.</value>
        public int Supbookpos
        {
            get{return supbookpos;}
            set { this.supbookpos = value; }
        }

        /// <summary>
        /// Gets or sets the externsheet pos.
        /// </summary>
        /// <value>The externsheet pos.</value>
        public int ExternsheetPos
        {
            get { return externsheetPos; }
            set { this.externsheetPos = value; }
        }

    }
}