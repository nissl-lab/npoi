
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


namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;

    /**
     * Title:        Footer Record 
     * Description:  Specifies the footer for a sheet
     * REFERENCE:  PG 317 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Shawn Laubach (slaubach at apache dot org) Modified 3/14/02
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class FooterRecord
       : HeaderFooterBase
    {
        public const short sid = 0x15;
        public FooterRecord(string text):base(text)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FooterRecord"/> class.
        /// </summary>
        /// <param name="in1">the RecordInputstream to Read the record from</param>
        public FooterRecord(RecordInputStream in1):base(in1)
        {
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FOOTER]\n");
            buffer.Append("    .footer         = ").Append(this.Text)
                .Append("\n");
            buffer.Append("[/FOOTER]\n");
            return buffer.ToString();
        }


        /// <summary>
        /// </summary>
        /// <value></value>
        /// return the non static version of the id for this record.
        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            return new FooterRecord(this.Text);
        }
    }
}