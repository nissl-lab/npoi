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
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.XPath;
using System.Xml;
using System.Xml.Schema;

namespace NPOI.Util
{
    public class DocumentHelper
    {
        private DocumentHelper()
        {

        }

        public static XPathDocument ReadDocument(Stream stream)
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationFlags = XmlSchemaValidationFlags.None;
            settings.ValidationType = ValidationType.None;
            settings.XmlResolver = null;
            settings.ProhibitDtd = true;
            //settings.ConformanceLevel = ConformanceLevel.Document;
            XmlReader xr = XmlReader.Create(stream, settings);
            
            XPathDocument xpathdoc = new XPathDocument(xr);
            return xpathdoc;
        }

        public static XmlDocument LoadDocument(Stream stream)
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationFlags = XmlSchemaValidationFlags.None;
            settings.ValidationType = ValidationType.Schema;
            settings.XmlResolver = null;
            settings.ProhibitDtd = true; 
            settings.ConformanceLevel = ConformanceLevel.Auto;
            settings.IgnoreProcessingInstructions = true;
            try
            {
                XmlReader xr = XmlReader.Create(stream, settings);

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.XmlResolver = null;
                xmlDoc.PreserveWhitespace = true;
                xmlDoc.Load(xr);

                return xmlDoc;
            }
            catch (XmlException)
            {
                //try to load using xml string, see TestExternalEntities.TestFile
                stream.Position = 0;
                using(StreamReader sr = new StreamReader(stream))
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.XmlResolver = null;
                    xmlDoc.PreserveWhitespace = true;
                    xmlDoc.LoadXml(sr.ReadToEnd());
                    return xmlDoc;
                }
            }
        }
    }
}
