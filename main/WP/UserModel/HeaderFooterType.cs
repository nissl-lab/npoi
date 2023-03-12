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

namespace NPOI.WP.UserModel
{
    /**
     * @since POI v3.16 beta 1
     */
    public enum HeaderFooterType : int
    {
        /**
         * This is the default header or Footer, It is displayed on every page where
         * a more specific header or footer is not specified. It is always displayed
         * on ODD pages that are not the first page of the section.
         */
        DEFAULT = 2,

        /**
         * This is an even page header or footer, it is displayed on even pages that
         * are not the first page of the section.
         */
        EVEN = 1,

        /**
         * This is a first page header or footer It is displayed on the first page
         * of the section.
         */
        FIRST = 3,
    }
}
