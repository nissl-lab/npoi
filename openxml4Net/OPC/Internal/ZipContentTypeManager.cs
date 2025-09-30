using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.Util;

#pragma warning disable CA1835

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

        public override async Task<bool> SaveImplAsync(XmlDocument content, Stream outStream, CancellationToken cancellationToken = default)
        {
            ZipOutputStream zos = null;
            if(outStream is ZipOutputStream stream)
                zos = stream;
            else
                zos = new ZipOutputStream(outStream);

            ZipEntry partEntry = new ZipEntry(CONTENT_TYPES_PART_NAME);
            try
            {
                // Referenced in ZIP
                await zos.PutNextEntryAsync(partEntry, cancellationToken);
                // Saving data in the ZIP file

                await StreamHelper.SaveXmlInStreamAsync(content, outStream, cancellationToken);
                using MemoryStream ins = new MemoryStream();

                byte[] buff = new byte[ZipHelper.READ_WRITE_FILE_BUFFER_SIZE];
                while(true)
                {
                    int resultRead = await ins.ReadAsync(buff, 0, ZipHelper.READ_WRITE_FILE_BUFFER_SIZE, cancellationToken);
                    if(resultRead == 0)
                    {
                        // end of file reached
                        break;
                    }
                    else
                    {
                        await zos.WriteAsync(buff, 0, resultRead, cancellationToken);
                    }
                }

                await zos.CloseEntryAsync(cancellationToken);
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