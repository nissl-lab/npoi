using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.Util;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /**
     * Zip implementation of the ContentTypeManager.
     *
     * @author Julien Chable
     * @version 1.0
     * @see ContentTypeManager
     */
    public class ZipContentTypeManager : ContentTypeManager
    {
        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(ZipContentTypeManager));

        /**
         * Delegate constructor to the super constructor.
         *
         * @param in
         *            The input stream to parse to fill internal content type
         *            collections.
         * @throws InvalidFormatException
         *             If the content types part content is not valid.
         */
        public ZipContentTypeManager(Stream in1, OPCPackage pkg) : base(in1, pkg)
        {
        }


        public override bool SaveImpl(XmlDocument content, Stream out1)
        {
            ZipOutputStream zos = null;
            if(out1 is ZipOutputStream stream)
                zos = stream;
            else
                zos = new ZipOutputStream(out1);

            ZipEntry partEntry = new ZipEntry(CONTENT_TYPES_PART_NAME);
            try
            {
                // Referenced in ZIP
                zos.PutNextEntry(partEntry);
                // Saving data in the ZIP file

                StreamHelper.SaveXmlInStream(content, out1);
                using MemoryStream ins = new MemoryStream();

                byte[] buff = new byte[ZipHelper.READ_WRITE_FILE_BUFFER_SIZE];
                while(true)
                {
                    int resultRead = ins.Read(buff, 0, ZipHelper.READ_WRITE_FILE_BUFFER_SIZE);
                    if(resultRead == 0)
                    {
                        // end of file reached
                        break;
                    }
                    else
                    {
                        zos.Write(buff, 0, resultRead);
                    }
                }

                zos.CloseEntry();
            }
            catch(IOException ioe)
            {
                logger.Log(POILogger.ERROR, "Cannot write: " + CONTENT_TYPES_PART_NAME
                                                             + " in Zip !", ioe);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Saves the content types asynchronously to the specified output stream.
        /// </summary>
        /// <param name="content">The XML content to save</param>
        /// <param name="out1">The output stream</param>
        /// <param name="cancellationToken">Cancellation token to observe during the async operation</param>
        /// <returns>A task that represents the asynchronous save operation</returns>
        public override async Task<bool> SaveImplAsync(XmlDocument content, Stream out1, CancellationToken cancellationToken)
        {
            ZipOutputStream zos = null;
            if(out1 is ZipOutputStream stream)
                zos = stream;
            else
                zos = new ZipOutputStream(out1);
            ZipEntry partEntry = new ZipEntry(CONTENT_TYPES_PART_NAME);
            try
            {
                // Referenced in ZIP
                zos.PutNextEntry(partEntry);
                // Saving data in the ZIP file asynchronously
                await StreamHelper.SaveXmlInStreamAsync(content, out1, cancellationToken).ConfigureAwait(false);

                using MemoryStream ins = new MemoryStream();

                byte[] buff = new byte[ZipHelper.READ_WRITE_FILE_BUFFER_SIZE];
                while(true)
                {
                    cancellationToken.ThrowIfCancellationRequested();
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
                    int resultRead = await ins.ReadAsync(buff.AsMemory(0, ZipHelper.READ_WRITE_FILE_BUFFER_SIZE), cancellationToken).ConfigureAwait(false);
#else
                    int resultRead = await ins.ReadAsync(buff, 0, ZipHelper.READ_WRITE_FILE_BUFFER_SIZE, cancellationToken).ConfigureAwait(false);
#endif
                    if(resultRead == 0)
                    {
                        // end of file reached
                        break;
                    }
                    else
                    {
#if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
                        await zos.WriteAsync(buff.AsMemory(0, resultRead), cancellationToken).ConfigureAwait(false);
#else
                        await zos.WriteAsync(buff, 0, resultRead, cancellationToken).ConfigureAwait(false);
#endif
                    }
                }

                zos.CloseEntry();
            }
            catch(IOException ioe)
            {
                logger.Log(POILogger.ERROR, "Cannot write: " + CONTENT_TYPES_PART_NAME
                                                             + " in Zip !", ioe);
                return false;
            }

            return true;
        }
    }
}