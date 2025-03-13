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
    using NPOI.HSSF.Record;
    using NPOI.SS.UserModel;

    /// <summary>
    /// Represents an Excel hyperlink.
    /// </summary>
    /// <remarks>@author Yegor Kozlov (yegor at apache dot org)</remarks>
    public class HSSFHyperlink : IHyperlink
    {
        /**
         * Low-level record object that stores the actual hyperlink data
         */
        public HyperlinkRecord record;

        /**
         * If we Create a new hypelrink remember its type
         */
        protected HyperlinkType link_type;

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFHyperlink"/> class.
        /// </summary>
        /// <param name="type">The type of hyperlink to Create.</param>
        public HSSFHyperlink(HyperlinkType type)
        {
            this.link_type = type;
            record = new HyperlinkRecord();
            switch (type)
            {
                case HyperlinkType.Url:
                case HyperlinkType.Email:
                    record.CreateUrlLink();
                    break;
                case HyperlinkType.File:
                    record.CreateFileLink();
                    break;
                case HyperlinkType.Document:
                    record.CreateDocumentLink();
                    break;
                default:
                    throw new ArgumentException("Invalid type: " + type);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFHyperlink"/> class.
        /// </summary>
        /// <param name="record">The record.</param>
        public HSSFHyperlink(HyperlinkRecord record)
        {
            this.record = record;
            link_type = getType(record);
        }
        private HyperlinkType getType(HyperlinkRecord record)
        {
            HyperlinkType link_type;
            // Figure out the type
            if (record.IsFileLink)
            {
                link_type = HyperlinkType.File;
            }
            else if (record.IsDocumentLink)
            {
                link_type = HyperlinkType.Document;
            }
            else
            {
                if (record.Address != null &&
                        record.Address.StartsWith("mailto:"))
                {
                    link_type = HyperlinkType.Email;
                }
                else
                {
                    link_type = HyperlinkType.Url;
                }
            }
            return link_type;
        }

        public HSSFHyperlink(IHyperlink other)
        {
            if (other is HSSFHyperlink hlink)
            {
                record = hlink.record.Clone() as HyperlinkRecord;
                link_type = getType(record);
            }
            else
            {
                link_type = other.Type;
                record = new HyperlinkRecord();
                FirstRow = (other.FirstRow);
                FirstColumn = (other.FirstColumn);
                LastRow = (other.LastRow);
                LastColumn = (other.LastColumn);
            }
        }

        /// <summary>
        /// Gets or sets the row of the first cell that Contains the hyperlink
        /// </summary>
        /// <value>the 0-based row of the cell that Contains the hyperlink.</value>
        public int FirstRow
        {
            get { return record.FirstRow; }
            set { record.FirstRow=value; }
        }

        /// <summary>
        /// Gets or sets the row of the last cell that Contains the hyperlink
        /// </summary>
        /// <value>the 0-based row of the last cell that Contains the hyperlink</value>
        public int LastRow
        {
            get { return record.LastRow; }
            set { record.LastRow=value; }
        }


        /// <summary>
        /// Gets or sets the column of the first cell that Contains the hyperlink
        /// </summary>
        /// <value>the 0-based column of the first cell that Contains the hyperlink</value>
        public int FirstColumn
        {
            get { return record.FirstColumn; }
            set { record.FirstColumn=value; }
        }

        /// <summary>
        /// Gets or sets the column of the last cell that Contains the hyperlink
        /// </summary>
        /// <value>the 0-based column of the last cell that Contains the hyperlink</value>
        public int LastColumn
        {
            get { return record.LastColumn; }
            set { record.LastColumn=value; }
        }

        /// <summary>
        /// Gets or sets Hypelink Address. Depending on the hyperlink type it can be URL, e-mail, patrh to a file, etc.
        /// </summary>
        /// <value>the Address of this hyperlink</value>
        public String Address
        {
            get
            {
                return record.Address;
            }
            set { record.Address = value; }
        }

        /// <summary>
        /// Gets or sets the text mark.
        /// </summary>
        /// <value>The text mark.</value>
        public String TextMark
        {
            get { return record.TextMark; }
            set { record.TextMark = value; }
        }

        /// <summary>
        /// Gets or sets the short filename.
        /// </summary>
        /// <value>The short filename.</value>
        public String ShortFilename
        {
            get { return record.ShortFilename; }
            set { record.ShortFilename = value; }
        }

        /// <summary>
        /// Gets or sets the text label for this hyperlink
        /// </summary>
        /// <value>text to Display</value>
        public String Label
        {
            get
            {
                return record.Label;
            }
            set 
            {
                record.Label=value;
            }
        }

        /// <summary>
        /// Gets the type of this hyperlink
        /// </summary>
        /// <value>the type of this hyperlink</value>
        public HyperlinkType Type
        {
            get { return (HyperlinkType)link_type; }
        }

        /**
         * @return whether the objects have the same HyperlinkRecord
         */
        public override bool Equals(Object other)
        {
            if (this == other) return true;
            if (other is not HSSFHyperlink otherLink) return false;
            return record == otherLink.record;
        }


        public override int GetHashCode()
        {
            return record.GetHashCode();
        }
    }
}