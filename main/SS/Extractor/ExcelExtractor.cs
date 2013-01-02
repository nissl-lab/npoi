/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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
namespace NPOI.SS.Extractor
{
    using System;

    /**
     * Common interface for Excel text extractors, covering
     *  HSSF and XSSF
     */
    public interface ExcelExtractor
    {
        /**
         * Should sheet names be included? Default is true
         */
        void SetIncludeSheetNames(bool includeSheetNames);

        /**
         * Should we return the formula itself, and not
         *  the result it produces? Default is false
         */
        void SetFormulasNotResults(bool formulasNotResults);

        /**
         * Should cell comments be included? Default is false
         */
        void SetIncludeCellComments(bool includeCellComments);

        /**
         * Retreives the text contents of the file
         */
        String Text { get; }
    }

}