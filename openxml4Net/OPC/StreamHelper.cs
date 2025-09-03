using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
            //XmlWriter writer = new XmlTextWriter(outStream,Encoding.UTF8);
            xmlContent.WriteContentTo(writer);
            writer.Flush();
        }

        /// <summary>
        /// Saves the XML document to the specified stream asynchronously.
        /// </summary>
        /// <param name="xmlContent">The XML document to save</param>
        /// <param name="outStream">The stream to write to</param>
        /// <param name="cancellationToken">Cancellation token to observe during the async operation</param>
        /// <returns>A task that represents the asynchronous save operation</returns>
        public static async Task SaveXmlInStreamAsync(XmlDocument xmlContent, Stream outStream, CancellationToken cancellationToken = default)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.OmitXmlDeclaration = false;
            settings.Indent = false;
            settings.Async = true; // Enable async writing
            
            using (XmlWriter writer = XmlWriter.Create(outStream, settings))
            {
                await writer.WriteRawAsync(xmlContent.OuterXml).ConfigureAwait(false);
                await writer.FlushAsync().ConfigureAwait(false);
            }
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

        /// <summary>
        /// Copy the input stream into the output stream asynchronously.
        /// </summary>
        /// <param name="inStream">The source stream</param>
        /// <param name="outStream">The destination stream</param>
        /// <param name="cancellationToken">Cancellation token to observe during the async operation</param>
        /// <returns>A task that represents the asynchronous copy operation</returns>
        public static async Task CopyStreamAsync(Stream inStream, Stream outStream, CancellationToken cancellationToken = default)
        {
            byte[] buffer = new byte[1024];
            int bytesRead = 0;
            while ((bytesRead = await inStream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
            {
                await outStream.WriteAsync(buffer, 0, bytesRead, cancellationToken).ConfigureAwait(false);
            }
        }
    }

}
