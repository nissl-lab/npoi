﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Threading.Tasks;
using System.Threading;

namespace NPOI.OpenXml4Net.OPC
{
    public class StreamHelper
    {

        private StreamHelper()
        {
            // Do nothing
        }

        /**
         * Turning the DOM4j object in the specified output stream.
         *
         * @param xmlContent
         *            The XML document.
         * @param outStream
         *            The Stream in which the XML document will be written.
         * @return <b>true</b> if the xml is successfully written in the stream,
         *         else <b>false</b>.
         */
        public static void SaveXmlInStream(XmlDocument xmlContent,
                Stream outStream)
        {
            //XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlContent.NameTable);
            //nsmgr.AddNamespace("", "");
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.OmitXmlDeclaration = false;
            // don't indent xml documents, the indent will cause errors in calculating the xml signature
            // because of different handling of linebreaks in Windows/Unix
            // see https://stackoverflow.com/questions/36063375
            settings.Indent = false;
            XmlWriter writer = XmlTextWriter.Create(outStream,settings);
            xmlContent.WriteContentTo(writer);
            writer.Flush();
        }

        public static async Task SaveXmlInStreamAsync(XmlDocument xmlContent,Stream outStream, CancellationToken cancellationToken = default)
        {
            //XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlContent.NameTable);
            //nsmgr.AddNamespace("", "");
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.OmitXmlDeclaration = false;
            // don't indent xml documents, the indent will cause errors in calculating the xml signature
            // because of different handling of linebreaks in Windows/Unix
            // see https://stackoverflow.com/questions/36063375
            settings.Indent = false;
            XmlWriter writer = XmlTextWriter.Create(outStream,settings);
            xmlContent.WriteContentTo(writer);
            await writer.FlushAsync();
        }

        /**
         * Copy the input stream into the output stream.
         *
         * @param inStream
         *            The source stream.
         * @param outStream
         *            The destination stream.
         * @return <b>true</b> if the operation succeed, else return <b>false</b>.
         */
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
