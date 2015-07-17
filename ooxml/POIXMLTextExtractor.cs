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

namespace NPOI
{
    using System;
    using NPOI.OpenXml4Net.OPC;

    public abstract class POIXMLTextExtractor : POITextExtractor
    {
        /** The POIXMLDocument that's open */
        private POIXMLDocument _document;

        /**
         * Creates a new text extractor for the given document
         */
        public POIXMLTextExtractor(POIXMLDocument document)
            : base((POIDocument)null)
        {


            _document = document;
        }

        /**
         * Returns the core document properties
         */
        public CoreProperties GetCoreProperties()
        {
            return _document.GetProperties().CoreProperties;
        }
        /**
         * Returns the extended document properties
         */
        public ExtendedProperties GetExtendedProperties()
        {
            return _document.GetProperties().ExtendedProperties;
        }
        /**
         * Returns the custom document properties
         */
        public CustomProperties GetCustomProperties()
        {
            return _document.GetProperties().CustomProperties;
        }

        /**
         * Returns opened document
         */
        public POIXMLDocument Document
        {
            get{
                return _document;
            }
        }

        /**
         * Returns the opened OPCPackage that Contains the document
         */
        public OPCPackage Package
        {
            get
            {
                return _document.Package;
            }
        }

        /**
         * Returns an OOXML properties text extractor for the
         *  document properties metadata, such as title and author.
         */
        public override POITextExtractor MetadataTextExtractor
        {
            get
            {
                return new POIXMLPropertiesTextExtractor(_document);
            }
        }

        public override void Close()
        {
            // e.g. XSSFEventBaseExcelExtractor passes a null-document
            if (_document != null)
            {
                OPCPackage pkg = _document.Package;
                if (pkg != null)
                {
                    pkg.Revert();
                }
            }
            base.Close();
        }
    }

}





