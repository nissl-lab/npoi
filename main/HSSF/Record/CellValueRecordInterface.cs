
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
        

/*
 * CellValueRecordInterface.java
 *
 * Created on October 2, 2001, 8:27 PM
 */
namespace NPOI.HSSF.Record
{

    /**
     * The cell value record interface Is implemented by all classes of type Record that
     * contain cell values.  It allows the containing sheet to move through them and Compare
     * them.
     *
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     *
     * @see org.apache.poi.hssf.model.Sheet
     * @see org.apache.poi.hssf.record.Record
     * @see org.apache.poi.hssf.record.RecordFactory
     */

    public interface CellValueRecordInterface
    {

        /**
         * Get the row this cell occurs on
         *
         * @return the row
         */

        //short Row;
        int Row { get; set; }

        /**
         * Get the column this cell defines within the row
         *
         * @return the column
         */

        int Column { get; set; }


        short XFIndex { get; set; }
    }
}