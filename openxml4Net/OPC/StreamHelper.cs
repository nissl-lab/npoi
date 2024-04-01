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
using System.Xml;
using System.IO;

namespace NPOI.OpenXml4Net.OPC
{
    public class StreamHelper
    {

        private StreamHelper()
        {
            // Do nothing
        }

        /// <summary>
        /// Turning the DOM4j object in the specified output stream.
        /// </summary>
        /// <param name="xmlContent">
        /// The XML document.
        /// </param>
        /// <param name="outStream">
        /// The Stream in which the XML document will be written.
        /// </param>
        /// <returns><b>true</b> if the xml is successfully written in the stream,
        /// else <b>false</b>.
        /// </returns>
        public static void SaveXmlInStream(XmlDocument xmlContent,
                Stream outStream)
        {
            //XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlContent.NameTable);
            //nsmgr.AddNamespace("", "");
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.OmitXmlDeclaration = false;
            XmlWriter writer = XmlTextWriter.Create(outStream,settings);
            //XmlWriter writer = new XmlTextWriter(outStream,Encoding.UTF8);
            xmlContent.WriteContentTo(writer);
            writer.Flush();
        }

        /// <summary>
        /// Copy the input stream into the output stream.
        /// </summary>
        /// <param name="inStream">
        /// The source stream.
        /// </param>
        /// <param name="outStream">
        /// The destination stream.
        /// </param>
        /// <returns><b>true</b> if the operation succeed, else return <b>false</b>.</returns>
        public static void CopyStream(Stream inStream, Stream outStream)
        {
            byte[] buffer = new byte[1024];
            int bytesRead = 0;
            int totalRead = 0;
            while ((bytesRead = inStream.Read(buffer, 0, buffer.Length)) > 0)
            {
                outStream.Write(buffer, 0, bytesRead);
                totalRead += bytesRead;
            }
        }
    }

}
