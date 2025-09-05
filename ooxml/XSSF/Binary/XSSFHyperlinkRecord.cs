/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XSSF.Binary
{
    using NPOI.SS.Util;
    using NPOI.Util;

    /// <summary>
    /// <para>
    /// This is a read only record that maintains information about
    /// a hyperlink.  In OOXML land, this information has to be merged
    /// from 1) the sheet's .rels to Get the url and 2) from After the
    /// sheet data in they hyperlink section.
    /// </para>
    /// <para>
    /// The <see cref="Display"/> is often empty and should be filled from
    /// the contents of the anchor cell.
    /// </para>
    /// </summary>
    /// @since 3.16-beta3
    public class XSSFHyperlinkRecord
    {

        private CellRangeAddress cellRangeAddress;
        private String relId;
        private String location;
        private String toolTip;
        private String display;

        public XSSFHyperlinkRecord(CellRangeAddress cellRangeAddress, String relId, String location, String toolTip,
            String display)
        {
            this.cellRangeAddress = cellRangeAddress;
            this.relId = relId;
            this.location = location;
            this.toolTip = toolTip;
            this.display = display;
        }



        public CellRangeAddress CellRangeAddress => cellRangeAddress;

        public String RelId => relId;

        public String Location
        {
            get => location;
            set => location = value;
        }

        public String ToolTip
        {
            get => toolTip;
            set => toolTip = value;
        }

        public String Display
        {
            get => display;
            set => display = value;
        }

        public override bool Equals(object o)
        {
            if(this == o) return true;
            if(o == null || GetType() != o.GetType()) return false;

            XSSFHyperlinkRecord that = (XSSFHyperlinkRecord) o;

            if(cellRangeAddress != null
                   ? !cellRangeAddress.Equals(that.cellRangeAddress)
                   : that.cellRangeAddress != null)
                return false;
            if(relId != null ? !relId.Equals(that.relId) : that.relId != null) return false;
            if(location != null ? !location.Equals(that.location) : that.location != null) return false;
            if(toolTip != null ? !toolTip.Equals(that.toolTip) : that.toolTip != null) return false;
            return display != null ? display.Equals(that.display) : that.display == null;
        }

        public override int GetHashCode()
        {
            int result = cellRangeAddress != null ? cellRangeAddress.GetHashCode() : 0;
            result = 31 * result + (relId != null ? relId.GetHashCode() : 0);
            result = 31 * result + (location != null ? location.GetHashCode() : 0);
            result = 31 * result + (toolTip != null ? toolTip.GetHashCode() : 0);
            result = 31 * result + (display != null ? display.GetHashCode() : 0);
            return result;
        }

        public override String ToString()
        {
            return "XSSFHyperlinkRecord{" +
                   "cellRangeAddress=" + cellRangeAddress +
                   ", relId='" + relId + '\'' +
                   ", location='" + location + '\'' +
                   ", toolTip='" + toolTip + '\'' +
                   ", display='" + display + '\'' +
                   '}';
        }
    }
}

