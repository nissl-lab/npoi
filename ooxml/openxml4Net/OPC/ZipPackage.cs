using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.OpenXml4Net.OPC.Internal.Marshallers;
using NPOI.OpenXml4Net.Exceptions;
using Util=NPOI.OpenXml4Net.Util;
using NPOI.OpenXml4Net.Util;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.Util;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Physical zip package.
     *
     * @author Julien Chable
     */
    public class ZipPackage : OPCPackage
    {

        private static POILogger logger = POILogFactory.GetLogger(typeof(ZipPackage));

        /**
         * Zip archive, as either a file on disk,
         *  or a stream
         */
        private Util.ZipEntrySource zipArchive;
        bool isStream = false;  // whether the file is passed in with stream, no means passed in with string path

        /**
         * Constructor. Creates a new ZipPackage.
         */
        public ZipPackage()
            : base(defaultPackageAccess)
        {
            this.zipArchive = null;
            try
            {
                this.contentTypeManager = new ZipContentTypeManager(null, this);
            }
            catch (InvalidFormatException e) { }
        }

        /**
         * Constructor. <b>Operation not supported.</b>
         *
         * @param in
         *            Zip input stream to load.
         * @param access
         * @throws ArgumentException
         *             If the specified input stream not an instance of
         *             ZipInputStream.
         */
        public ZipPackage(Stream filestream, PackageAccess access)
            : base(access)
        {
            isStream = true;
            this.zipArchive = new ZipInputStreamZipEntrySource(new ZipInputStream(filestream));
        }

        /**
         * Constructor. Opens a Zip based Open XML document.
         *
         * @param path
         *            The path of the file to open or create.
         * @param access
         *            The package access mode.
         * @throws InvalidFormatException
         *             If the content type part parsing encounters an error.
         */
        public ZipPackage(String path, PackageAccess access)
            : base(access)
        {
            ZipFile zipFile = null;

            try
            {
                zipFile = ZipHelper.OpenZipFile(path);
            }
            catch (IOException e)
            {
                throw new InvalidOperationException(
                      "Can't open the specified file: '" + path + "'", e);
            }

            this.zipArchive = new ZipFileZipEntrySource(zipFile);
        }

        /**
         * Constructor. Opens a Zip based Open XML document.
         *
         * @param file
         *            The file to open or create.
         * @param access
         *            The package access mode.
         * @throws InvalidFormatException
         *             If the content type part parsing encounters an error.
         */
        public ZipPackage(FileInfo file, PackageAccess access)
            : base(access)
        {


            ZipFile zipFile = null;

            try
            {
                zipFile = ZipHelper.OpenZipFile(file);
            }
            catch (IOException e)
            {
                throw new InvalidOperationException(
                      "Can't open the specified file: '" + file + "'", e);
            }

            this.zipArchive = new ZipFileZipEntrySource(zipFile);
        }

        /**
         * Retrieves the parts from this package. We assume that the package has not
         * been yet inspect to retrieve all the parts, this method will open the
         * archive and look for all parts contain inside it. If the package part
         * list is not empty, it will be emptied.
         *
         * @return All parts contain in this package.
         * @throws InvalidFormatException
         *             Throws if the package is not valid.
         */

        protected override PackagePart[] GetPartsImpl(){
        if (this.partList == null) {
            // The package has just been created, we create an empty part
            // list.
            this.partList = new PackagePartCollection();
        }

        if (this.zipArchive == null)
        {
            PackagePart[] pp = new PackagePart[this.partList.Values.Count];
            this.partList.Values.CopyTo(pp, 0);
            return pp;
        }
        // First we need to parse the content type part
        IEnumerator entries = this.zipArchive.Entries;
        while (entries.MoveNext()) {
            ZipEntry entry = (ZipEntry)entries.Current;
            if (entry.Name.ToLower().Equals(
                    ContentTypeManager.CONTENT_TYPES_PART_NAME.ToLower()))
            {
                try {
                    this.contentTypeManager = new ZipContentTypeManager(
                            ZipArchive.GetInputStream(entry), this);
                } catch (IOException e) {
                    throw new InvalidFormatException(e.Message);
                }
                break;
            }
        }

        // At this point, we should have loaded the content type part
        if (this.contentTypeManager == null) {
            throw new InvalidFormatException(
                    "Package should contain a content type part [M1.13]");
        }

        // Now create all the relationships
        // (Need to create relationships before other
        //  parts, otherwise we might create a part before
        //  its relationship exists, and then it won't tie up)
        entries = this.zipArchive.Entries;
        while (entries.MoveNext()) {
            ZipEntry entry = (ZipEntry)entries.Current;
            PackagePartName partName = BuildPartName(entry);
            if(partName == null) continue;

            // Only proceed for Relationships at this stage
            String contentType = contentTypeManager.GetContentType(partName);
            if (contentType != null && contentType.Equals(ContentTypes.RELATIONSHIPS_PART)) {
                try {
                    partList[partName]= new ZipPackagePart(this, entry,
                        partName, contentType);
                } catch (InvalidOperationException e) {
                    throw new InvalidFormatException(e.Message);
                }
            }
        }

        // Then we can go through all the other parts
        entries = this.zipArchive.Entries;
        while (entries.MoveNext()) {
            ZipEntry entry = entries.Current as ZipEntry;
            PackagePartName partName = BuildPartName(entry);
            if(partName == null) continue;

            String contentType = contentTypeManager
                    .GetContentType(partName);
            if (contentType != null && contentType.Equals(ContentTypes.RELATIONSHIPS_PART)) {
                // Already handled
            }
            else if (contentType != null) 
            {
                try {
                    partList[partName]= new ZipPackagePart(this, entry,
                            partName, contentType);
                } catch (InvalidOperationException e) {
                    throw new InvalidFormatException(e.Message);
                }
            } else {
                throw new InvalidFormatException(
                        "The part "
                                + partName.URI.OriginalString
                                + " does not have any content type ! Rule: Package require content types when retrieving a part from a package. [M.1.14]");
            }
        }
            ZipPackagePart[] returnArray =new ZipPackagePart[partList.Count];
            partList.Values.CopyTo(returnArray,0);
        return returnArray;
    }

        /**
         * Builds a PackagePartName for the given ZipEntry,
         *  or null if it's the content types / invalid part
         */
        private PackagePartName BuildPartName(ZipEntry entry)
        {
            try
            {
                // We get an error when we parse [Content_Types].xml
                // because it's not a valid URI.
                if (entry.Name.ToLower().Equals(
                        ContentTypeManager.CONTENT_TYPES_PART_NAME.ToLower()))
                {
                    return null;
                }
                return PackagingUriHelper.CreatePartName(ZipHelper
                        .GetOPCNameFromZipItemName(entry.Name));
            }
            catch (Exception)
            {
                // We assume we can continue, even in degraded mode ...
                //logger.log(POILogger.WARN,"Entry "
                //                + entry.getName()
                //                + " is not valid, so this part won't be add to the package.");
                return null;
            }
        }

        /**
         * Create a new MemoryPackagePart from the specified URI and content type
         *
         *
         * aram partName The part URI.
         *
         * @param contentType
         *            The part content type.
         * @return The newly created zip package part, else <b>null</b>.
         */

        protected override PackagePart CreatePartImpl(PackagePartName partName,
                String contentType, bool loadRelationships)
        {
            if (contentType == null)
                throw new ArgumentException("contentType");

            if (partName == null)
                throw new ArgumentException("partName");

            try
            {
                return new MemoryPackagePart(this, partName, contentType,
                        loadRelationships);
            }
            catch (InvalidFormatException)
            {
                // TODO - don't use system.err.  Is it valid to return null when this exception occurs?
                //System.err.println(e);
                return null;
            }
        }

        /**
         * Delete a part from the package
         *
         * @throws ArgumentException
         *             Throws if the part URI is nulll or invalid.
         */

        protected override void RemovePartImpl(PackagePartName partName)
        {
            if (partName == null)
                throw new ArgumentException("partUri");
        }

        /**
         * Flush the package. Do nothing.
         */

        protected override void FlushImpl()
        {
            // Do nothing
        }

        /**
         * Close and save the package.
         *
         * @see #close()
         */

        protected override void CloseImpl()
        {
            // Flush the package
            Flush();

            // Save the content
            if (this.originalPackagePath != null
                    && !"".Equals(this.originalPackagePath))
            {
                if (File.Exists(this.originalPackagePath))
                {
                    // Case of a package previously open

                    string tempfilePath=GenerateTempFileName(FileHelper
                                    .GetDirectory(this.originalPackagePath));

                    FileInfo fi=NPOI.Util.TempFile.CreateTempFile(
                            tempfilePath, ".tmp");

                    // Save the final package to a temporary file
                    try
                    {
                        FileStream fs = File.OpenWrite(fi.FullName);
                        Save(fs);
                        if(zipArchive!=null)
                            this.zipArchive.Close(); // Close the zip archive to be
                        fs.Close();
                        // able to delete it
                        FileHelper.CopyFile(fi.FullName, this.originalPackagePath);
                    }
                    finally
                    {
                        // Either the save operation succeed or not, we delete the
                        // temporary file
                        File.Delete(tempfilePath);
                        
                            logger
                                    .Log(POILogger.WARN, "The temporary file: '"
                                            + tempfilePath
                                            + "' cannot be deleted ! Make sure that no other application use it.");
                        
                    }
                }
                else
                {
                    throw new InvalidOperationException(
                            "Can't close a package not previously open with the open() method !");
                }
            }
        }

        /**
         * Create a unique identifier to be use as a temp file name.
         *
         * @return A unique identifier use to be use as a temp file name.
         */
        private String GenerateTempFileName(string directory)
        {
            FileInfo tmpFilename = null ;
            string path = null;
            do
            {
                path = directory + "\\"
            + "OpenXml4Net" + System.DateTime.Now.Ticks;

                tmpFilename = new FileInfo(path);
            } while (File.Exists(path));

            return tmpFilename.Name; //FileHelper.getFilename(tmpFilename.Name);
        }

        /**
         * Close the package without saving the document. Discard all the changes
         * made to this package.
         */

        protected override void RevertImpl()
        {
            try
            {
                if (this.zipArchive != null)
                    this.zipArchive.Close();
            }
            catch (IOException)
            {
                // Do nothing, user dont have to know
            }
        }

        /**
         * Implement the getPart() method to retrieve a part from its URI in the
         * current package
         *
         *
         * @see #getPart(PackageRelationship)
         */

        protected override PackagePart GetPartImpl(PackagePartName partName)
        {
            if (partList.ContainsKey(partName))
            {
                return partList[partName];
            }
            return null;
        }

        /**
         * Save this package into the specified stream
         *
         *
         * @param outputStream
         *            The stream use to save this package.
         *
         * @see #save(OutputStream)
         */

        protected override void SaveImpl(Stream outputStream)
        {
            // Check that the document was open in write mode
            ThrowExceptionIfReadOnly();
            ZipOutputStream zos = null;

            try
            {
                if (!(outputStream is ZipOutputStream))
                    zos = new ZipOutputStream(outputStream);
                else
                    zos = (ZipOutputStream)outputStream;

                zos.UseZip64 = UseZip64.Off;
                // If the core properties part does not exist in the part list,
                // we save it as well
                if (this.GetPartsByRelationshipType(PackageRelationshipTypes.CORE_PROPERTIES).Count == 0 &&
                this.GetPartsByRelationshipType(PackageRelationshipTypes.CORE_PROPERTIES_ECMA376).Count == 0)
                {
                    logger.Log(POILogger.DEBUG, "Save core properties part");

                    // Ensure that core properties are added if missing
                    GetPackageProperties();

                    // Add core properties to part list ...
                    AddPackagePart(this.packageProperties);

                    // ... and to add its relationship ...
                    this.relationships.AddRelationship(this.packageProperties
                            .PartName.URI, TargetMode.Internal,
                            PackageRelationshipTypes.CORE_PROPERTIES, null);
                    // ... and the content if it has not been added yet.
                    if (!this.contentTypeManager
                            .IsContentTypeRegister(ContentTypes.CORE_PROPERTIES_PART))
                    {
                        this.contentTypeManager.AddContentType(
                                this.packageProperties.PartName,
                                ContentTypes.CORE_PROPERTIES_PART);
                    }
                }

                // Save package relationships part.
                logger.Log(POILogger.DEBUG,"Save package relationships");
                ZipPartMarshaller.MarshallRelationshipPart(this.Relationships,
                        PackagingUriHelper.PACKAGE_RELATIONSHIPS_ROOT_PART_NAME,
                        zos);

                // Save content type part.
                logger.Log(POILogger.DEBUG,"Save content types part");
                this.contentTypeManager.Save(zos);

                // Save parts.
                foreach (PackagePart part in GetParts())
                {
                    // If the part is a relationship part, we don't save it, it's
                    // the source part that will do the job.
                    if (part.IsRelationshipPart)
                        continue;

                    logger.Log(POILogger.DEBUG,"Save part '"
                            + ZipHelper.GetZipItemNameFromOPCName(part
                                    .PartName.Name) + "'");
                    if (partMarshallers.ContainsKey(part._contentType))
                    {
                        PartMarshaller marshaller = partMarshallers[part._contentType];

                        if (!marshaller.Marshall(part, zos))
                        {
                            throw new OpenXml4NetException(
                                    "The part "
                                            + part.PartName.URI
                                            + " fail to be saved in the stream with marshaller "
                                            + marshaller);
                        }
                    }
                    else
                    {
                        if (!defaultPartMarshaller.Marshall(part, zos))
                            throw new OpenXml4NetException(
                                    "The part "
                                            + part.PartName.URI
                                            + " fail to be saved in the stream with marshaller "
                                            + defaultPartMarshaller);
                    }
                }
                //Finishes writing the contents of the ZIP output stream without closing the underlying stream.
                if (isStream)
                    zos.Finish();   //instead of use zos.Close, it will close the stream
                else
                    zos.Close();
            }
            catch (Exception e)
            {
                logger.Log(POILogger.ERROR, "fail to save: an error occurs while saving the package : "
                                + e.Message);
            }
        }

        /**
         * Get the zip archive
         *
         * @return The zip archive.
         */
        public Util.ZipEntrySource ZipArchive
        {
            get
            {
                return zipArchive;
            }
        }
    }

}
