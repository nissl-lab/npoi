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
using NPOI.OpenXml4Net.OPC;
namespace NPOI
{


    /**
     * Defines a factory API that enables sub-classes to create instances of <code>POIXMLDocumentPart</code>
     *
     * @author Yegor Kozlov
     */
    public abstract class POIXMLFactory
    {

        /**
         * Create a POIXMLDocumentPart from existing namespace part and relation. This method is called
         * from {@link POIXMLDocument#load(POIXMLFactory)} when parsing a document
         *
         * @param parent parent part
         * @param rel   the namespace part relationship
         * @param part  the PackagePart representing the Created instance
         * @return A new instance of a POIXMLDocumentPart.
         */
        public abstract POIXMLDocumentPart CreateDocumentPart(POIXMLDocumentPart parent, PackageRelationship rel, PackagePart part);

        /**
         * Create a new POIXMLDocumentPart using the supplied descriptor. This method is used when Adding new parts
         * to a document, for example, when Adding a sheet to a workbook, slide to a presentation, etc.
         *
         * @param descriptor  describes the object to create
         * @return A new instance of a POIXMLDocumentPart.
         */
        public abstract POIXMLDocumentPart CreateDocumentPart(POIXMLRelation descriptor);
    }


}




