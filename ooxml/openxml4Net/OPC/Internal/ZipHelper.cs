using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

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
                return null;
            }

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
            if (!File.Exists(path))
            {
                return null;
            }

            return new ZipFile(File.OpenRead(path));
        }

    }
}