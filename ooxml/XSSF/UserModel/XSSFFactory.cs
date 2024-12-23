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

using System;
namespace NPOI.XSSF.UserModel
{

    /**
     * Instantiates sub-classes of POIXMLDocumentPart depending on their relationship type
     *
     * @author Yegor Kozlov
     */
    public sealed class XSSFFactory : POIXMLFactory
    {
        private XSSFFactory()
        {

        }

        private static readonly XSSFFactory inst = new XSSFFactory();

        public static XSSFFactory GetInstance()
        {
            return inst;
        }

        /**
         * @since POI 3.14-Beta1
         */
        protected override POIXMLRelation GetDescriptor(String relationshipType)
        {
            return XSSFRelation.GetInstance(relationshipType);
        }
    }
}