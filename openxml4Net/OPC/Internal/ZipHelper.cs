using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.POIFS.Common;
using NPOI.Util;
using NPOI.POIFS.Storage;
using NPOI.Openxml4Net.Exceptions;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    public class ZipHelper
    {

        /**
         * Forward slash use to convert part name between OPC and zip item naming
         * conventions.
         */
        private static String FORWARD_SLASH = "/";

        /**
         * Buffer to read data from file. Use big buffer to improve performaces. the
         * InputStream class is reading only 8192 bytes per read call (default value
         * set by sun)
         */
        public static int READ_WRITE_FILE_BUFFER_SIZE = 8192;

        /**
         * Prevent this class to be instancied.
         */
        private ZipHelper()
        {
            // Do nothing
        }

        /**
         * Retrieve the zip entry of the core properties part.
         *
         * @throws OpenXml4NetException
         *             Throws if internal error occurs.
         */
        public static ZipEntry GetCorePropertiesZipEntry(ZipPackage pkg)
        {
            PackageRelationship corePropsRel = pkg.GetRelationshipsByType(
                    PackageRelationshipTypes.CORE_PROPERTIES).GetRelationship(0);

            if (corePropsRel == null)
                return null;

            ZipEntry ze = new ZipEntry(corePropsRel.TargetUri.OriginalString);
            return ze;
        }

        /**
         * Retrieve the Zip entry of the content types part.
         */
        public static ZipEntry GetContentTypeZipEntry(ZipPackage pkg)
        {
            IEnumerator entries = pkg.ZipArchive.Entries;
            // Enumerate through the Zip entries until we find the one named
            // '[Content_Types].xml'.
            while (entries.MoveNext())
            {
                ZipEntry entry = (ZipEntry)entries.Current;
                if (entry.Name.Equals(
                        ContentTypeManager.CONTENT_TYPES_PART_NAME))
                    return entry;
            }
            return null;
        }

        /**
         * Convert a zip name into an OPC name by adding a leading forward slash to
         * the specified item name.
         *
         * @param zipItemName
         *            Zip item name to convert.
         * @return An OPC compliant name.
         */
        public static String GetOPCNameFromZipItemName(String zipItemName)
        {
            if (zipItemName == null)
                throw new ArgumentException("zipItemName");
            if (zipItemName.StartsWith(FORWARD_SLASH))
                return zipItemName;
            else
                return FORWARD_SLASH + zipItemName;
        }

        /**
         * Convert an OPC item name into a zip item name by removing any leading
         * forward slash if it exist.
         *
         * @param opcItemName
         *            The OPC item name to convert.
         * @return A zip item name without any leading slashes.
         */
        public static String GetZipItemNameFromOPCName(String opcItemName)
        {
            if (opcItemName == null)
                throw new ArgumentException("opcItemName");

            String retVal = opcItemName;
            while (retVal.StartsWith(FORWARD_SLASH))
                retVal = retVal.Substring(1);
            return retVal;
        }

        /**
         * Convert an OPC item name into a zip URI by removing any leading forward
         * slash if it exist.
         *
         * @param opcItemName
         *            The OPC item name to convert.
         * @return A zip URI without any leading slashes.
         */
        public static Uri GetZipURIFromOPCName(String opcItemName)
        {
            if (opcItemName == null)
                throw new ArgumentException("opcItemName");

            String retVal = opcItemName;
            while (retVal.StartsWith(FORWARD_SLASH))
                retVal = retVal.Substring(1);
            try
            {
                return PackagingUriHelper.ParseUri(retVal, UriKind.RelativeOrAbsolute);
            }
            catch (UriFormatException)
            {
                return null;
            }
        }

        /**
         * Verifies that the given stream starts with a Zip structure.
         * 
         * Warning - this will consume the first few bytes of the stream,
         *  you should push-back or reset the stream after use!
         */
        public static void VerifyZipHeader(InputStream stream)
        {
            // Grab the first 8 bytes
            byte[] data = new byte[8];
            IOUtils.ReadFully(stream, data);

            // OLE2?
            long signature = LittleEndian.GetLong(data);
            if (signature == HeaderBlockConstants._signature)
            {
                throw new OLE2NotOfficeXmlFileException(
                    "The supplied data appears to be in the OLE2 Format. " +
                    "You are calling the part of POI that deals with OOXML " +
                    "(Office Open XML) Documents. You need to call a different " +
                    "part of POI to process this data (eg HSSF instead of XSSF)");
            }

            // Raw XML?
            byte[] RAW_XML_FILE_HEADER = POIFSConstants.RAW_XML_FILE_HEADER;
            if (data[0] == RAW_XML_FILE_HEADER[0] &&
                data[1] == RAW_XML_FILE_HEADER[1] &&
                data[2] == RAW_XML_FILE_HEADER[2] &&
                data[3] == RAW_XML_FILE_HEADER[3] &&
                data[4] == RAW_XML_FILE_HEADER[4])
            {
                throw new NotOfficeXmlFileException(
                    "The supplied data appears to be a raw XML file. " +
                    "Formats such as Office 2003 XML are not supported");
            }

            // Don't check for a Zip header, as to maintain backwards
            //  compatibility we need to let them seek over junk at the
            //  start before beginning processing.

            // Put things back
            if (stream is PushbackInputStream)
            {
                ((PushbackInputStream)stream).Unread(data);
            }
            else if (stream.MarkSupported())
            {
                stream.Reset();
            }
            else
            {
                // Oh dear... I hope you know what you're doing!
            }
        }

        private static InputStream PrepareToCheckHeader(InputStream stream)
        {
            if (stream is PushbackInputStream)
            {
                return stream;
            }
            if (stream.MarkSupported())
            {
                stream.Mark(8);
                return stream;
            }
            return new PushbackInputStream(stream, 8);
        }
        // TODO: ZipSecureFile
        /**
         * Opens the specified stream as a secure zip
         *
         * @param stream
         *            The stream to open.
         * @return The zip stream freshly open.
         */
        //public static ThresholdInputStream OpenZipStream(Stream stream)
        //{
        //    // Peek at the first few bytes to sanity check
        //    InputStream checkedStream = prepareToCheckHeader(stream);
        //    verifyZipHeader(checkedStream);

        //    // Open as a proper zip stream
        //    InputStream zis = new ZipInputStream(checkedStream);
        //    return ZipSecureFile.addThreshold(zis);
        //}

        public static ZipInputStream OpenZipStream(Stream stream)
        {
            // TODO: ZipSecureFile
            //InputStream zis = new ZipInputStream(stream);
            //ThresholdInputStream tis = ZipSecureFile.AddThreshold(zis);
            //return tis;
            return new ZipInputStream(stream);
        }
        /**
        * Opens the specified file as a zip, or returns null if no such file exists
        *
        * @param file
        *            The file to open.
        * @return The zip archive freshly open.
        */
        public static ZipFile OpenZipFile(FileInfo file)
        {
            if (!file.Exists)
            {
                throw new FileNotFoundException("File does not exist");
            }
            //if (file.isDirectory())
            //{
            //    throw new IOException("File is a directory");
            //}

            // Peek at the first few bytes to sanity check
            FileInputStream input = new FileInputStream(file.OpenRead());
            try
            {
                VerifyZipHeader(input);
            }
            finally
            {
                input.Close();
            }
            // TODO: ZipSecureFile
            //// Open as a proper zip file
            //return new ZipSecureFile(file);
            return new ZipFile(File.OpenRead(file.FullName));
        }
        /**
         * Retrieve and open a zip file with the specified path.
         *
         * @param path
         *            The file path.
         * @return The zip archive freshly open.
         */
        public static ZipFile OpenZipFile(String path)
        {
            return OpenZipFile(new FileInfo(path));
        }

    }
}