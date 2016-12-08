using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC.Internal.Marshallers;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.OpenXml4Net.OPC.Internal.Unmarshallers;
using NPOI.Util;
using System.Text.RegularExpressions;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Represents a container that can store multiple data objects.
     * 
     * @author Julien Chable, CDubet
     * @version 0.1
     */
    public abstract class OPCPackage : RelationshipSource
    {

        /**
         * Logger.
         */
        private static POILogger logger = POILogFactory.GetLogger(typeof(OPCPackage));

        /**
         * Default package access.
         */
        protected static PackageAccess defaultPackageAccess = PackageAccess.READ_WRITE;

        /**
         * Package access.
         */
        private PackageAccess packageAccess;

        /**
         * Package parts collection.
         */
        protected PackagePartCollection partList;

        /**
         * Package relationships.
         */
        protected PackageRelationshipCollection relationships;

        /**
         * Part marshallers by content type.
         */
        protected SortedList<ContentType, PartMarshaller> partMarshallers;

        /**
         * Default part marshaller.
         */
        protected PartMarshaller defaultPartMarshaller;

        /**
         * Part unmarshallers by content type.
         */
        protected SortedList<ContentType, PartUnmarshaller> partUnmarshallers;

        /**
         * Core package properties.
         */
        protected PackagePropertiesPart packageProperties;

        /**
         * Manage parts content types of this package.
         */
        protected ContentTypeManager contentTypeManager;

        /**
         * Flag if a modification is done to the document.
         */
        protected bool isDirty = false;

        /**
         * File path of this package.
         */
        protected String originalPackagePath;

        /**
         * Output stream for writing this package.
         */
        protected Stream output;

        /**
         * Constructor.
         * 
         * @param access
         *            Package access.
         */
        public OPCPackage(PackageAccess access)
        {
            if (GetType() != typeof(ZipPackage))
            {
                throw new ArgumentException("PackageBase may not be subclassed");
            }
            Init();
            this.packageAccess = access;
        }

        /**
         * Initialize the package instance.
         */
        private void Init()
        {
            this.partMarshallers = new SortedList<ContentType, PartMarshaller>(5);
            this.partUnmarshallers = new SortedList<ContentType, PartUnmarshaller>(2);

            try
            {
                // Add 'default' unmarshaller
                this.partUnmarshallers.Add(new ContentType(
                        ContentTypes.CORE_PROPERTIES_PART),
                        new PackagePropertiesUnmarshaller());

                // Add default marshaller
                this.defaultPartMarshaller = new DefaultMarshaller();
                // TODO Delocalize specialized marshallers
                this.partMarshallers.Add(new ContentType(
                        ContentTypes.CORE_PROPERTIES_PART),
                        new ZipPackagePropertiesMarshaller());
            }
            catch (InvalidFormatException)
            {
                // Should never happen
                throw new OpenXml4NetException(
                        "Package.init() : this exception should never happen, if you read this message please send a mail to the developers team.");
            }
        }


        /**
         * Open a package with read/write permission.
         * 
         * @param path
         *            The document path.
         * @return A Package object, else <b>null</b>.
         * @throws InvalidFormatException
         *             If the specified file doesn't exist, and a parsing error
         *             occur.
         */
        public static OPCPackage Open(String path)
        {
            return Open(path, defaultPackageAccess);
        }
        /**
        * Open a package with read/write permission.
        *
        * @param file
        *            The file to open.
        * @return A Package object, else <b>null</b>.
        * @throws InvalidFormatException
        *             If the specified file doesn't exist, and a parsing error
        *             occur.
        */
        public static OPCPackage Open(FileInfo file)
        {
            return Open(file, defaultPackageAccess);
        }
        /**
         * Open a package.
         * 
         * @param path
         *            The document path.
         * @param access
         *            PackageBase access.
         * @return A PackageBase object, else <b>null</b>.
         * @throws InvalidFormatException
         *             If the specified file doesn't exist, and a parsing error
         *             occur.
         */
        public static OPCPackage Open(String path, PackageAccess access)
        {
            if (path == null || "".Equals(path.Trim()))
                throw new ArgumentException("'path' must be given");

            if (new DirectoryInfo(path).Exists)
                throw new ArgumentException("path must not be a directory");


            OPCPackage pack = new ZipPackage(path, access);
            if (pack.partList == null && access != PackageAccess.WRITE)
            {
                pack.GetParts();
            }
            pack.originalPackagePath = new DirectoryInfo(path).FullName;
            return pack;
        }
        /**
        * Open a package.
        *
        * @param file
        *            The file to open.
        * @param access
        *            PackageBase access.
        * @return A PackageBase object, else <b>null</b>.
        * @throws InvalidFormatException
        *             If the specified file doesn't exist, and a parsing error
        *             occur.
        */
        public static OPCPackage Open(FileInfo file, PackageAccess access)
        {
            if (file == null)
                throw new ArgumentException("'file' must be given");
            if (new DirectoryInfo(file.FullName).Exists)
                throw new ArgumentException("file must not be a directory");

            OPCPackage pack = new ZipPackage(file, access);
            if (pack.partList == null && access != PackageAccess.WRITE)
            {
                pack.GetParts();
            }
            pack.originalPackagePath = file.FullName;
            return pack;
        }
        /**
         * Open a package.
         * 
         * Note - uses quite a bit more memory than {@link #open(String)}, which
         * doesn't need to hold the whole zip file in memory, and can take advantage
         * of native methods
         * 
         * @param in
         *            The InputStream to read the package from
         * @return A PackageBase object
         */
        public static OPCPackage Open(Stream in1)
        {
            OPCPackage pack = new ZipPackage(in1, PackageAccess.READ_WRITE);
            if (pack.partList == null)
            {
                pack.GetParts();
            }
            return pack;
        }

        /**
         * Opens a package if it exists, else it Creates one.
         * 
         * @param file
         *            The file to open or to Create.
         * @return A newly Created package if the specified file does not exist,
         *         else the package extract from the file.
         * @throws InvalidFormatException
         *             Throws if the specified file exist and is not valid.
         */
        public static OPCPackage OpenOrCreate(string path)
        {
            OPCPackage retPackage = null;
            if (File.Exists(path))
            {
                retPackage = Open(path);
            }
            else
            {
                retPackage = Create(path);
            }
            return retPackage;
        }

        /**
         * Creates a new package.
         * 
         * @param file
         *            Path of the document.
         * @return A newly Created PackageBase ready to use.
         */
        public static OPCPackage Create(string path)
        {
            if (new DirectoryInfo(path).Exists)
                throw new ArgumentException("file");

            if (File.Exists(path))
            {
                throw new InvalidOperationException(
                        "This package (or file) already exists : use the open() method or delete the file.");
            }

            // Creates a new package
            OPCPackage pkg = null;
            pkg = new ZipPackage();
            pkg.originalPackagePath = (new FileInfo(path)).Name;

            ConfigurePackage(pkg);
            return pkg;
        }

        public static OPCPackage Create(Stream output)
        {
            OPCPackage pkg = null;
            pkg = new ZipPackage();
            pkg.originalPackagePath = null;
            pkg.output = output;

            ConfigurePackage(pkg);
            return pkg;
        }

        /**
         * Configure the package.
         * 
         * @param pkg
         */
        private static void ConfigurePackage(OPCPackage pkg)
        {
            try
            {
                // Content type manager
                pkg.contentTypeManager = new ZipContentTypeManager(null, pkg);
                // Add default content types for .xml and .rels
                pkg.contentTypeManager.AddContentType(
                                PackagingUriHelper.CreatePartName(PackagingUriHelper.PACKAGE_RELATIONSHIPS_ROOT_URI),
                                ContentTypes.RELATIONSHIPS_PART);
                pkg.contentTypeManager.AddContentType(PackagingUriHelper
                                .CreatePartName("/default.xml"),
                                ContentTypes.PLAIN_OLD_XML);

                // Init some PackageBase properties
                pkg.packageProperties = new PackagePropertiesPart(pkg,
                        PackagingUriHelper.CORE_PROPERTIES_PART_NAME);
                pkg.packageProperties.SetCreatorProperty("Generated by OpenXml4Net");
                pkg.packageProperties.SetCreatedProperty(new Nullable<DateTime>(DateTime.Now));
            }
            catch (InvalidFormatException)
            {
                // Should never happen
                throw;
            }
        }

        /**
         * Flush the package : save all.
         * 
         * @see #close()
         */
        public void Flush()
        {
            ThrowExceptionIfReadOnly();

            if (this.packageProperties != null)
            {
                this.packageProperties.Flush();
            }

            this.FlushImpl();
        }

        /**
         * Close the package and save its content.
         * 
         * @throws IOException
         *             If an IO exception occur during the saving process.
         */
        public void Close()
        {
            if (this.packageAccess == PackageAccess.READ)
            {
                logger
                        .Log(POILogger.WARN, "The close() method is intended to SAVE a package. This package is open in READ ONLY mode, use the revert() method instead !");
                Revert();
                return;
            }
            if (this.contentTypeManager == null)
            {
                logger.Log(POILogger.WARN,
                        "Unable to call close() on a package that hasn't been fully opened yet");
                return;
            }
            // Save the content
            //ReentrantReadWriteLock l = new ReentrantReadWriteLock();
            try
            {
                //l.writeLock().lock();
                if (this.originalPackagePath != null
                        && !"".Equals(this.originalPackagePath.Trim()))
                {
                    //File targetFile = new File(this.originalPackagePath);
                    if (!File.Exists(this.originalPackagePath))
                    {
                        //|| !(this.originalPackagePath
                        //		.Equals(targetFile.GetAbsolutePath(),StringComparison.InvariantCultureIgnoreCase))) {
                        // Case of a package Created from scratch
                        Save(originalPackagePath);
                    }
                    else
                    {
                        CloseImpl();
                    }
                }
                else if (this.output != null)
                {
                    Save(this.output);
                }
            }
            finally
            {
                //l.writeLock().unlock();
            }

            // Clear
            this.contentTypeManager.ClearAll();

            // Call the garbage collector
            //Runtime.GetRuntime().gc();
        }

        /**
         * Close the package WITHOUT saving its content. Reinitialize this package
         * and cancel all changes done to it.
         */
        public void Revert()
        {
            RevertImpl();
        }

        /**
         * Add a thumbnail to the package. This method is provided to make easier
         * the addition of a thumbnail in a package. You can do the same work by
         * using the traditionnal relationship and part mechanism.
         * 
         * @param path
         *            The full path to the image file.
         */
        public void AddThumbnail(String path)
        {
            // Check parameter
            if ("".Equals(path))
                throw new ArgumentException("path");

            // Get the filename from the path
            String filename = path
                    .Substring(path.LastIndexOf('\\') + 1);

            // Create the thumbnail part name
            String contentType = ContentTypes
                    .GetContentTypeFromFileExtension(filename);
            PackagePartName thumbnailPartName = null;
            try
            {
                thumbnailPartName = PackagingUriHelper.CreatePartName("/docProps/"
                        + filename);
            }
            catch (InvalidFormatException)
            {
                try
                {
                    thumbnailPartName = PackagingUriHelper
                            .CreatePartName("/docProps/thumbnail"
                                    + path.Substring(path.LastIndexOf(".") + 1));
                }
                catch (InvalidFormatException)
                {
                    throw new InvalidOperationException(
                            "Can't add a thumbnail file named '" + filename + "'");
                }
            }

            // Check if part already exist
            if (this.GetPart(thumbnailPartName) != null)
                throw new InvalidOperationException(
                        "You already add a thumbnail named '" + filename + "'");

            // Add the thumbnail part to this package.
            PackagePart thumbnailPart = this.CreatePart(thumbnailPartName,
                    contentType, false);

            // Add the relationship between the package and the thumbnail part
            this.AddRelationship(thumbnailPartName, TargetMode.Internal,
                    PackageRelationshipTypes.THUMBNAIL);

            // Copy file data to the newly Created part
            StreamHelper.CopyStream(new FileStream(path, FileMode.Open), thumbnailPart
                    .GetOutputStream());
        }

        /**
         * Throws an exception if the package access mode is in read only mode
         * (PackageAccess.Read).
         * 
         * @throws InvalidOperationException
         *             Throws if a writing operation is done on a read only package.
         * @see org.apache.poi.OpenXml4Net.opc.PackageAccess
         */
        internal void ThrowExceptionIfReadOnly()
        {
            if (packageAccess == PackageAccess.READ)
                throw new InvalidOperationException(
                        "Operation not allowed, document open in read only mode!");
        }

        /**
         * Throws an exception if the package access mode is in write only mode
         * (PackageAccess.Write). This method is call when other methods need write
         * right.
         * 
         * @throws InvalidOperationException
         *             Throws if a read operation is done on a write only package.
         * @see org.apache.poi.OpenXml4Net.opc.PackageAccess
         */
        internal void ThrowExceptionIfWriteOnly()
        {
            if (packageAccess == PackageAccess.WRITE)
                throw new InvalidOperationException(
                        "Operation not allowed, document open in write only mode!");
        }

        /**
         * Retrieves or Creates if none exists, core package property part.
         * 
         * @return The PackageProperties part of this package.
         */
        public PackageProperties GetPackageProperties()
        {
            this.ThrowExceptionIfWriteOnly();
            // If no properties part has been found then we Create one
            if (this.packageProperties == null)
            {
                this.packageProperties = new PackagePropertiesPart(this,
                        PackagingUriHelper.CORE_PROPERTIES_PART_NAME);
            }
            return this.packageProperties;
        }

        public bool PartExists(Uri uri)
        {
            if (uri.IsAbsoluteUri)
                return false;

            PackagePart pp = GetPartImpl(new PackagePartName(uri.OriginalString, true));
            return pp != null;
        }
        public PackagePart GetPart(Uri uri)
        {
            ThrowExceptionIfWriteOnly();
            PackagePartName partName = new PackagePartName(uri.ToString(), true);
            if (partName == null)
                throw new ArgumentException("PartName");

            // If the partlist is null, then we parse the package.
            if (partList == null)
            {
                try
                {
                    GetParts();
                }
                catch (InvalidFormatException)
                {
                    return null;
                }
            }
            return GetPartImpl(partName);
        }
        /**
         * Retrieve a part identified by its name.
         * 
         * @param PartName
         *            Part name of the part to retrieve.
         * @return The part with the specified name, else <code>null</code>.
         */
        public PackagePart GetPart(PackagePartName partName)
        {
            ThrowExceptionIfWriteOnly();

            if (partName == null)
                throw new ArgumentException("PartName");

            // If the partlist is null, then we parse the package.
            if (partList == null)
            {
                try
                {
                    GetParts();
                }
                catch (InvalidFormatException)
                {
                    return null;
                }
            }
            return GetPartImpl(partName);
        }

        /**
         * Retrieve parts by content type.
         * 
         * @param contentType
         *            The content type criteria.
         * @return All part associated to the specified content type.
         */
        public List<PackagePart> GetPartsByContentType(String contentType)
        {
            List<PackagePart> retArr = new List<PackagePart>();
            foreach (PackagePart part in partList.Values)
            {
                if (part.ContentType.Equals(contentType))
                    retArr.Add(part);
            }
            retArr.Sort();
            return retArr;
        }

        /**
         * Retrieve parts by relationship type.
         * 
         * @param relationshipType
         *            Relationship type.
         * @return All parts which are the target of a relationship with the
         *         specified type, if the method can't retrieve relationships from
         *         the package, then return <code>null</code>.
         */
        public List<PackagePart> GetPartsByRelationshipType(
                String relationshipType)
        {
            if (relationshipType == null)
                throw new ArgumentException("relationshipType");
            List<PackagePart> retArr = new List<PackagePart>();
            foreach (PackageRelationship rel in GetRelationshipsByType(relationshipType))
            {
                PackagePart part = GetPart(rel);
                if (part != null)
                {
                    retArr.Add(part);
                }
            }
            retArr.Sort();
            return retArr;
        }
        /**
         * Retrieve parts by name
         *
         * @param namePattern
         *            The pattern for matching the names
         * @return All parts associated to the specified content type, sorted
         * in alphanumerically by the part-name
         */
        public List<PackagePart> GetPartsByName(Regex namePattern)
        {
            if (namePattern == null)
            {
                throw new ArgumentException("name pattern must not be null");
            }
            List<PackagePart> result = new List<PackagePart>();
            foreach (PackagePart part in partList.Values)
            {
                PackagePartName partName = part.PartName;
                String name = partName.Name;
                if (namePattern.IsMatch(name))
                    result.Add(part);
            }
            result.Sort();
            return result;
        }

        /**
         * Get the target part from the specified relationship.
         * 
         * @param partRel
         *            The part relationship uses to retrieve the part.
         */
        public PackagePart GetPart(PackageRelationship partRel)
        {
            PackagePart retPart = null;
            EnsureRelationships();
            foreach (PackageRelationship rel in relationships)
            {
                if (rel.RelationshipType.Equals(partRel.RelationshipType))
                {
                    try
                    {
                        retPart = GetPart(PackagingUriHelper.CreatePartName(rel
                                .TargetUri));
                    }
                    catch (InvalidFormatException)
                    {
                        continue;
                    }
                    break;
                }
            }
            return retPart;
        }

        /**
         * Load the parts of the archive if it has not been done yet. The
         * relationships of each part are not loaded.
         * Note - Rule M4.1 states that there may only ever be one Core
         *  Properties Part, but Office produced files will sometimes
         *  have multiple! As Office ignores all but the first, we relax
         *  Compliance with Rule M4.1, and ignore all others silently too. 
         * @return All this package's parts.
         */
        public List<PackagePart> GetParts()
        {
            ThrowExceptionIfWriteOnly();

            // If the part list is null, we parse the package to retrieve all parts.
            if (partList == null)
            {
                /* Variables use to validate OPC Compliance */

                // Check rule M4.1 -> A format consumer shall consider more than
                // one core properties relationship for a package to be an error
                // (We just log it and move on, as real files break this!)
                bool hasCorePropertiesPart = false;
                bool needCorePropertiesPart = true;

                PackagePart[] parts = this.GetPartsImpl();
                this.partList = new PackagePartCollection();
                foreach (PackagePart part in parts)
                {
                    bool pnFound = false;
                    foreach (PackagePartName pn in partList.Keys)
                    {
                        if (part.PartName.Name.StartsWith(pn.Name))
                        {
                            pnFound = true;
                            break;
                        }
                    }

                    if (pnFound)
                        throw new InvalidFormatException(
                                "A part with the name '"
                                        + part.PartName +
                                        "' already exist : Packages shall not contain equivalent " +
                                    "part names and package implementers shall neither create " +
                                    "nor recognize packages with equivalent part names. [M1.12]");

                    // Check OPC compliance rule M4.1
                    if (part.ContentType.Equals(
                            ContentTypes.CORE_PROPERTIES_PART))
                    {
                        if (!hasCorePropertiesPart)
                            hasCorePropertiesPart = true;
                        else
                            Console.WriteLine(
                                    "OPC Compliance error [M4.1]: there is more than one core properties relationship in the package ! " +
                                    "POI will use only the first, but other software may reject this file.");
                    }



                    if (partUnmarshallers.ContainsKey(part._contentType))
                    {
                        PartUnmarshaller partUnmarshaller = partUnmarshallers[part._contentType];
                        UnmarshallContext context = new UnmarshallContext(this,
                                part.PartName);
                        try
                        {
                            PackagePart unmarshallPart = partUnmarshaller
                                    .Unmarshall(context, part.GetInputStream());
                            partList[unmarshallPart.PartName] = unmarshallPart;

                            // Core properties case-- use first CoreProperties part we come across
                            // and ignore any subsequent ones
                            if (unmarshallPart is PackagePropertiesPart &&
                                    hasCorePropertiesPart &&
                                    needCorePropertiesPart)
                            {
                                this.packageProperties = (PackagePropertiesPart)unmarshallPart;
                                needCorePropertiesPart = false;
                            }
                        }
                        catch (IOException)
                        {
                            logger.Log(POILogger.WARN, "Unmarshall operation : IOException for "
                                    + part.PartName);
                            continue;
                        }
                        catch (InvalidOperationException invoe)
                        {
                            throw new InvalidFormatException(invoe.Message);
                        }
                    }
                    else
                    {
                        try
                        {
                            partList[part.PartName] = part;
                        }
                        catch (InvalidOperationException e)
                        {
                            throw new InvalidFormatException(e.Message);
                        }
                    }
                }
            }
            return new List<PackagePart>(partList.Values);
        }
        public PackagePart CreatePart(Uri partName, String contentType)
        {
            return this.CreatePart(new PackagePartName(partName.OriginalString, true), contentType, true);
        }
        /**
         * Create and Add a part, with the specified name and content type, to the
         * package.
         * 
         * @param PartName
         *            Part name.
         * @param contentType
         *            Part content type.
         * @return The newly Created part.
         * @throws InvalidFormatException
         *             If rule M1.12 is not verified : Packages shall not contain
         *             equivalent part names and package implementers shall neither
         *             Create nor recognize packages with equivalent part names.
         * @see #CreatePartImpl(PackagePartName, String, bool) 
         */
        public PackagePart CreatePart(PackagePartName partName, String contentType)
        {
            return this.CreatePart(partName, contentType, true);
        }

        /**
         * Create and Add a part, with the specified name and content type, to the
         * package. For general purpose, prefer the overload version of this method
         * without the 'loadRelationships' parameter.
         * 
         * @param PartName
         *            Part name.
         * @param contentType
         *            Part content type.
         * @param loadRelationships
         *            Specify if the existing relationship part, if any, logically
         *            associated to the newly Created part will be loaded.
         * @return The newly Created part.
         * @throws InvalidFormatException
         *             If rule M1.12 is not verified : Packages shall not contain
         *             equivalent part names and package implementers shall neither
         *             Create nor recognize packages with equivalent part names.
         * @see {@link#CreatePartImpl(URI, String)}
         */
        public PackagePart CreatePart(PackagePartName partName, String contentType,
                bool loadRelationships)
        {
            ThrowExceptionIfReadOnly();
            if (partName == null)
            {
                throw new ArgumentException("PartName");
            }

            if (contentType == null || contentType == "")
            {
                throw new ArgumentException("contentType");
            }
            bool pnFound = false;
            bool pnDeleted = false;
            foreach (PackagePartName pn in partList.Keys)
            {
                if (partName.Name.StartsWith(pn.Name))
                {
                    pnFound = true;
                    if (partList[pn].IsDeleted)
                    {
                        pnDeleted = true;
                    }
                    break;
                }
            }
            // Check if the specified part name already exists
            if (pnFound
                    && !pnDeleted)
            {
                throw new PartAlreadyExistsException(
                        "A part with the name '" + partName.Name + "'" +
                        " already exists : Packages shall not contain equivalent part names and package" +
                        " implementers shall neither create nor recognize packages with equivalent part names. [M1.12]");
            }

            /* Check OPC compliance */

            // Rule [M4.1]: The format designer shall specify and the format producer
            // producer
            // shall Create at most one core properties relationship for a package.
            // A format consumer shall consider more than one core properties
            // relationship for a package to be an error. If present, the
            // relationship shall target the Core Properties part.
            // Note - POI will read files with more than one Core Properties, which
            //  Office sometimes produces, but is strict on generation
            if (contentType == ContentTypes.CORE_PROPERTIES_PART)
            {
                if (this.packageProperties != null)
                    throw new InvalidOperationException(
                            "OPC Compliance error [M4.1]: you try to add more than one core properties relationship in the package !");
            }

            /* End check OPC compliance */

            PackagePart part = this.CreatePartImpl(partName, contentType,
                    loadRelationships);
            this.contentTypeManager.AddContentType(partName, contentType);
            this.partList[partName] = part;
            this.isDirty = true;
            return part;
        }

        /**
         * Add a part to the package.
         * 
         * @param PartName
         *            Part name of the part to Create.
         * @param contentType
         *            type associated with the file
         * @param content
         *            the contents to Add. In order to have faster operation in
         *            document merge, the data are stored in memory not on a hard
         *            disk
         * 
         * @return The new part.
         * @see #CreatePart(PackagePartName, String)
         */
        public PackagePart CreatePart(PackagePartName partName, String contentType,
                MemoryStream content)
        {
            PackagePart AddedPart = this.CreatePart(partName, contentType);
            if (AddedPart == null)
            {
                return null;
            }
            // Extract the zip entry content to put it in the part content
            if (content != null)
            {
                try
                {
                    Stream partOutput = AddedPart.GetOutputStream();
                    if (partOutput == null)
                    {
                        return null;
                    }

                    partOutput.Write(content.ToArray(), 0, (int)content.Length);
                    partOutput.Close();

                }
                catch (IOException)
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
            return AddedPart;
        }

        /**
         * Add the specified part to the package. If a part already exists in the
         * package with the same name as the one specified, then we replace the old
         * part by the specified part.
         * 
         * @param part
         *            The part to Add (or replace).
         * @return The part Added to the package, the same as the one specified.
         * @throws InvalidFormatException
         *             If rule M1.12 is not verified : Packages shall not contain
         *             equivalent part names and package implementers shall neither
         *             Create nor recognize packages with equivalent part names.
         */
        protected PackagePart AddPackagePart(PackagePart part)
        {
            ThrowExceptionIfReadOnly();
            if (part == null)
            {
                throw new ArgumentException("part");
            }

            if (partList.ContainsKey(part.PartName))
            {
                if (!partList[part.PartName].IsDeleted)
                {
                    throw new InvalidOperationException(
                            "A part with the name '"
                                    + part.PartName.Name
                                    + "' already exists : Packages shall not contain equivalent part names and package implementers shall neither Create nor recognize packages with equivalent part names. [M1.12]");
                }
                // If the specified partis flagged as deleted, we make it
                // available
                part.IsDeleted = false;
                // and delete the old part to replace it thereafeter
                this.partList.Remove(part.PartName);
            }
            this.partList[part.PartName] = part;
            this.isDirty = true;
            return part;
        }

        /**
         * Remove the specified part in this package. If this part is relationship
         * part, then delete all relationships in the source part.
         * 
         * @param part
         *            The part to Remove. If <code>null</code>, skip the action.
         * @see #RemovePart(PackagePartName)
         */
        public void RemovePart(PackagePart part)
        {
            if (part != null)
            {
                RemovePart(part.PartName);
            }
        }

        /**
         * Remove a part in this package. If this part is relationship part, then
         * delete all relationships in the source part.
         * 
         * @param PartName
         *            The part name of the part to Remove.
         */
        public void RemovePart(PackagePartName PartName)
        {
            ThrowExceptionIfReadOnly();
            if (PartName == null || !this.ContainPart(PartName))
                throw new ArgumentException("PartName");

            // Delete the specified part from the package.
            if (this.partList.ContainsKey(PartName))
            {
                this.partList[PartName].IsDeleted = (true);
                this.RemovePartImpl(PartName);
                this.partList.Remove(PartName);
            }
            else
            {
                this.RemovePartImpl(PartName);
            }

            // Delete content type
            this.contentTypeManager.RemoveContentType(PartName);

            // If this part is a relationship part, then delete all relationships of
            // the source part.
            if (PartName.IsRelationshipPartURI())
            {
                Uri sourceURI = PackagingUriHelper
                        .GetSourcePartUriFromRelationshipPartUri(PartName.URI);
                PackagePartName sourcePartName;
                try
                {
                    sourcePartName = PackagingUriHelper.CreatePartName(sourceURI);
                }
                catch (InvalidFormatException)
                {
                    logger
                            .Log(POILogger.ERROR, "Part name URI '"
                                    + sourceURI
                                    + "' is not valid ! This message is not intended to be displayed !");
                    return;
                }
                if (sourcePartName.URI.Equals(
                        PackagingUriHelper.PACKAGE_ROOT_URI))
                {
                    ClearRelationships();
                }
                else if (ContainPart(sourcePartName))
                {
                    PackagePart part = GetPart(sourcePartName);
                    if (part != null)
                        part.ClearRelationships();
                }
            }

            this.isDirty = true;
        }

        /**
         * Remove a part from this package as well as its relationship part, if one
         * exists, and all parts listed in the relationship part. Be aware that this
         * do not delete relationships which target the specified part.
         * 
         * @param PartName
         *            The name of the part to delete.
         * @throws InvalidFormatException
         *             Throws if the associated relationship part of the specified
         *             part is not valid.
         */
        public void RemovePartRecursive(PackagePartName PartName)
        {
            // Retrieves relationship part, if one exists
            PackagePart relPart = this.partList[PackagingUriHelper
                    .GetRelationshipPartName(PartName)];
            // Retrieves PackagePart object from the package
            PackagePart partToRemove = this.partList[PartName];

            if (relPart != null)
            {
                PackageRelationshipCollection partRels = new PackageRelationshipCollection(
                        partToRemove);
                foreach (PackageRelationship rel in partRels)
                {
                    PackagePartName PartNameToRemove = PackagingUriHelper
                            .CreatePartName(PackagingUriHelper.ResolvePartUri(rel
                                    .SourceUri, rel.TargetUri));
                    RemovePart(PartNameToRemove);
                }

                // Finally delete its relationship part if one exists
                this.RemovePart(relPart.PartName);
            }

            // Delete the specified part
            this.RemovePart(partToRemove.PartName);
        }
        public void DeletePart(Uri uri)
        {
            PackagePartName partName = new PackagePartName(uri.ToString(), true);
            if (partName == null)
                throw new ArgumentException("PartName");

            // Remove the part
            this.RemovePart(partName);
            // Remove the relationships part
            this.RemovePart(PackagingUriHelper.GetRelationshipPartName(partName));
        }
        /**
         * Delete the part with the specified name and its associated relationships
         * part if one exists. Prefer the use of this method to delete a part in the
         * package, compare to the Remove() methods that don't Remove associated
         * relationships part.
         * 
         * @param PartName
         *            Name of the part to delete
         */
        public void DeletePart(PackagePartName partName)
        {
            if (partName == null)
                throw new ArgumentException("PartName");

            // Remove the part
            this.RemovePart(partName);
            // Remove the relationships part
            this.RemovePart(PackagingUriHelper.GetRelationshipPartName(partName));
        }

        /**
         * Delete the part with the specified name and all part listed in its
         * associated relationships part if one exists. This process is recursively
         * apply to all parts in the relationships part of the specified part.
         * Prefer the use of this method to delete a part in the package, compare to
         * the Remove() methods that don't Remove associated relationships part.
         * 
         * @param PartName
         *            Name of the part to delete
         */
        public void DeletePartRecursive(PackagePartName partName)
        {
            if (partName == null || !this.ContainPart(partName))
                throw new ArgumentException("PartName");

            PackagePart partToDelete = this.GetPart(partName);
            // Remove the part
            this.RemovePart(partName);
            // Remove all relationship parts associated
            try
            {
                foreach (PackageRelationship relationship in partToDelete
                        .Relationships)
                {
                    PackagePartName targetPartName = PackagingUriHelper
                            .CreatePartName(PackagingUriHelper.ResolvePartUri(
                                    partName.URI, relationship.TargetUri));
                    this.DeletePartRecursive(targetPartName);
                }
            }
            catch (InvalidFormatException e)
            {
                logger.Log(POILogger.WARN, "An exception occurs while deleting part '"
                        + partName.Name
                        + "'. Some parts may remain in the package. - "
                        + e.Message);
                return;
            }
            // Remove the relationships part
            PackagePartName relationshipPartName = PackagingUriHelper
                    .GetRelationshipPartName(partName);
            if (relationshipPartName != null && ContainPart(relationshipPartName))
                this.RemovePart(relationshipPartName);
        }

        /**
         * Check if a part already exists in this package from its name.
         * 
         * @param PartName
         *            Part name to check.
         * @return <i>true</i> if the part is logically Added to this package, else
         *         <i>false</i>.
         */
        public bool ContainPart(PackagePartName partName)
        {
            return (this.GetPart(partName) != null);
        }

        /**
         * Add a relationship to the package (except relationships part).
         * 
         * Check rule M4.1 : The format designer shall specify and the format
         * producer shall Create at most one core properties relationship for a
         * package. A format consumer shall consider more than one core properties
         * relationship for a package to be an error. If present, the relationship
         * shall target the Core Properties part.
         * 
         * Check rule M1.25: The Relationships part shall not have relationships to
         * any other part. Package implementers shall enforce this requirement upon
         * the attempt to Create such a relationship and shall treat any such
         * relationship as invalid.
         * 
         * @param targetPartName
         *            Target part name.
         * @param targetMode
         *            Target mode, either Internal or External.
         * @param relationshipType
         *            Relationship type.
         * @param relID
         *            ID of the relationship.
         * @see PackageRelationshipTypes
         */
        public PackageRelationship AddRelationship(PackagePartName targetPartName,
                TargetMode targetMode, String relationshipType, String relID)
        {
            /* Check OPC compliance */

            // Check rule M4.1 : The format designer shall specify and the format
            // producer
            // shall Create at most one core properties relationship for a package.
            // A format consumer shall consider more than one core properties
            // relationship for a package to be an error. If present, the
            // relationship shall target the Core Properties part.
            if (relationshipType.Equals(PackageRelationshipTypes.CORE_PROPERTIES)
                    && this.packageProperties != null)
                throw new InvalidOperationException(
                        "OPC Compliance error [M4.1]: can't add another core properties part ! Use the built-in package method instead.");

            /*
             * Check rule M1.25: The Relationships part shall not have relationships
             * to any other part. Package implementers shall enforce this
             * requirement upon the attempt to Create such a relationship and shall
             * treat any such relationship as invalid.
             */
            if (targetPartName.IsRelationshipPartURI())
            {
                throw new InvalidOperationException(
                        "Rule M1.25: The Relationships part shall not have relationships to any other part.");
            }

            /* End OPC compliance */

            EnsureRelationships();
            PackageRelationship retRel = relationships.AddRelationship(
                    targetPartName.URI, targetMode, relationshipType, relID);
            this.isDirty = true;
            return retRel;
        }

        /**
         * Add a package relationship.
         * 
         * @param targetPartName
         *            Target part name.
         * @param targetMode
         *            Target mode, either Internal or External.
         * @param relationshipType
         *            Relationship type.
         * @see PackageRelationshipTypes
         */
        public PackageRelationship AddRelationship(PackagePartName targetPartName,
                TargetMode targetMode, String relationshipType)
        {
            return this.AddRelationship(targetPartName, targetMode,
                    relationshipType, null);
        }

        /**
         * Adds an external relationship to a part (except relationships part).
         * 
         * The targets of external relationships are not subject to the same
         * validity checks that internal ones are, as the contents is potentially
         * any file, URL or similar.
         * 
         * @param target
         *            External target of the relationship
         * @param relationshipType
         *            Type of relationship.
         * @return The newly Created and Added relationship
         * @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#AddExternalRelationship(java.lang.String,
         *      java.lang.String)
         */
        public PackageRelationship AddExternalRelationship(String target,
                String relationshipType)
        {
            return AddExternalRelationship(target, relationshipType, null);
        }

        /**
         * Adds an external relationship to a part (except relationships part).
         * 
         * The targets of external relationships are not subject to the same
         * validity checks that internal ones are, as the contents is potentially
         * any file, URL or similar.
         * 
         * @param target
         *            External target of the relationship
         * @param relationshipType
         *            Type of relationship.
         * @param id
         *            Relationship unique id.
         * @return The newly Created and Added relationship
         * @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#AddExternalRelationship(java.lang.String,
         *      java.lang.String)
         */
        public PackageRelationship AddExternalRelationship(String target,
                String relationshipType, String id)
        {
            if (target == null)
            {
                throw new ArgumentException("target");
            }
            if (relationshipType == null)
            {
                throw new ArgumentException("relationshipType");
            }

            Uri targetURI;
            try
            {
                targetURI = PackagingUriHelper.ParseUri(target,UriKind.Absolute);
            }
            catch (UriFormatException e)
            {
                throw new ArgumentException("Invalid target - " + e);
            }

            EnsureRelationships();
            PackageRelationship retRel = relationships.AddRelationship(targetURI,
                    TargetMode.External, relationshipType, id);
            this.isDirty = true;
            return retRel;
        }

        /**
         * Delete a relationship from this package.
         * 
         * @param id
         *            Id of the relationship to delete.
         */
        public void RemoveRelationship(String id)
        {
            if (relationships != null)
            {
                relationships.RemoveRelationship(id);
                this.isDirty = true;
            }
        }

        /**
         * Retrieves all package relationships.
         * 
         * @return All package relationships of this package.
         * @throws OpenXml4NetException
         * @see #GetRelationshipsHelper(String)
         */
        public PackageRelationshipCollection Relationships
        {
            get
            {
                return GetRelationshipsHelper(null);
            }
        }

        /**
         * Retrieves all relationships with the specified type.
         * 
         * @param relationshipType
         *            The filter specifying the relationship type.
         * @return All relationships with the specified relationship type.
         */
        public PackageRelationshipCollection GetRelationshipsByType(
                String relationshipType)
        {
            ThrowExceptionIfWriteOnly();
            if (relationshipType == null)
            {
                throw new ArgumentException("relationshipType");
            }
            return GetRelationshipsHelper(relationshipType);
        }

        /**
         * Retrieves all relationships with specified id (normally just ine because
         * a relationship id is supposed to be unique).
         * 
         * @param id
         *            Id of the wanted relationship.
         */
        private PackageRelationshipCollection GetRelationshipsHelper(String id)
        {
            ThrowExceptionIfWriteOnly();
            EnsureRelationships();
            return this.relationships.GetRelationships(id);
        }

        /**
         * Clear package relationships.
         */
        public void ClearRelationships()
        {
            if (relationships != null)
            {
                relationships.Clear();
                this.isDirty = true;
            }
        }

        /**
         * Ensure that the relationships collection is not null.
         */
        public void EnsureRelationships()
        {
            if (this.relationships == null)
            {
                try
                {
                    this.relationships = new PackageRelationshipCollection(this);
                }
                catch (InvalidFormatException)
                {
                    this.relationships = new PackageRelationshipCollection();
                }
            }
        }

        /**
         * @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#GetRelationship(java.lang.String)
         */
        public PackageRelationship GetRelationship(String id)
        {
            return this.relationships.GetRelationshipByID(id);
        }

        /**
         * @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#hasRelationships()
         */
        public bool HasRelationships
        {
            get
            {
                return (relationships.Size > 0);
            }
        }

        /**
         * @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#isRelationshipExists(org.apache.poi.OpenXml4Net.opc.PackageRelationship)
         */
        public bool IsRelationshipExists(PackageRelationship rel)
        {
            foreach (PackageRelationship r in this.Relationships)
            {
                if (r == rel)
                    return true;
            }
            return false;
        }

        /**
         * Add a marshaller.
         * 
         * @param contentType
         *            The content type to bind to the specified marshaller.
         * @param marshaller
         *            The marshaller to register with the specified content type.
         */
        public void AddMarshaller(String contentType, PartMarshaller marshaller)
        {
            try
            {
                partMarshallers[new ContentType(contentType)] = marshaller;
            }
            catch (InvalidFormatException e)
            {
                logger.Log(POILogger.WARN, "The specified content type is not valid: '"
                        + e.Message + "'. The marshaller will not be Added !");
            }
        }

        /**
         * Add an unmarshaller.
         * 
         * @param contentType
         *            The content type to bind to the specified unmarshaller.
         * @param unmarshaller
         *            The unmarshaller to register with the specified content type.
         */
        public void AddUnmarshaller(String contentType,
                PartUnmarshaller unmarshaller)
        {
            try
            {
                partUnmarshallers[new ContentType(contentType)] = unmarshaller;
            }
            catch (InvalidFormatException e)
            {
                logger.Log(POILogger.WARN, "The specified content type is not valid: '"
                        + e.Message
                        + "'. The unmarshaller will not be Added !");
            }
        }

        /**
         * Remove a marshaller by its content type.
         * 
         * @param contentType
         *            The content type associated with the marshaller to Remove.
         */
        public void RemoveMarshaller(String contentType)
        {
            partMarshallers.Remove(new ContentType(contentType));
        }

        /**
         * Remove an unmarshaller by its content type.
         * 
         * @param contentType
         *            The content type associated with the unmarshaller to Remove.
         */
        public void RemoveUnmarshaller(String contentType)
        {
            partUnmarshallers.Remove(new ContentType(contentType));
        }

        /* Accesseurs */

        /**
         * Get the package access mode.
         * 
         * @return the packageAccess The current package access.
         */
        public PackageAccess GetPackageAccess()
        {
            return packageAccess;
        }

        /**
         * Validates the package compliance with the OPC specifications.
         * 
         * @return <b>true</b> if the package is valid else <b>false</b>
         */
        public bool ValidatePackage(OPCPackage pkg)
        {
            throw new InvalidOperationException("Not implemented yet !!!");
        }

        /**
         * Save the document in the specified file.
         * 
         * @param targetFile
         *            Destination file.
         * @throws IOException
         *             Throws if an IO exception occur.
         * @see #save(OutputStream)
         */
        public void Save(string path)
        {
            if (path == null)
                throw new ArgumentException("targetFile");

            this.ThrowExceptionIfReadOnly();
            FileStream fos = null;
            try
            {
                fos = new FileStream(path, FileMode.OpenOrCreate);
            }
            catch (IOException e)
            {
                throw new IOException(e.Message, e);
            }
            try
            {
                this.Save(fos);
            }
            finally
            {
                fos.Close();
            }
        }

        /**
         * Save the document in the specified output stream.
         * 
         * @param outputStream
         *            The stream to save the package.
         * @see #saveImpl(OutputStream)
         */
        public void Save(Stream outputStream)
        {
            ThrowExceptionIfReadOnly();
            this.SaveImpl(outputStream);
        }

        /**
         * Core method to Create a package part. This method must be implemented by
         * the subclass.
         * 
         * @param PartName
         *            URI of the part to Create.
         * @param contentType
         *            Content type of the part to Create.
         * @return The newly Created package part.
         */
        protected abstract PackagePart CreatePartImpl(PackagePartName PartName,
                String contentType, bool loadRelationships);

        /**
         * Core method to delete a package part. This method must be implemented by
         * the subclass.
         * 
         * @param PartName
         *            The URI of the part to delete.
         */
        protected abstract void RemovePartImpl(PackagePartName PartName);

        /**
         * Flush the package but not save.
         */
        protected abstract void FlushImpl();

        /**
         * Close the package and cause a save of the package.
         * 
         */
        protected abstract void CloseImpl();

        /**
         * Close the package without saving the document. Discard all changes made
         * to this package.
         */
        protected abstract void RevertImpl();

        /**
         * Save the package into the specified output stream.
         * 
         * @param outputStream
         *            The output stream use to save this package.
         */
        protected abstract void SaveImpl(Stream outputStream);

        /**
         * Get the package part mapped to the specified URI.
         * 
         * @param PartName
         *            The URI of the part to retrieve.
         * @return The package part located by the specified URI, else <b>null</b>.
         */
        protected abstract PackagePart GetPartImpl(PackagePartName PartName);

        /**
         * Get all parts link to the package.
         * 
         * @return A list of the part owned by the package.
         */
        protected abstract PackagePart[] GetPartsImpl();
        /**
         * Replace a content type in this package.
         *
         * <p>
         *     A typical scneario to call this method is to rename a template file to the main format, e.g.
         *     ".dotx" to ".docx"
         *     ".dotm" to ".docm"
         *     ".xltx" to ".xlsx"
         *     ".xltm" to ".xlsm"
         *     ".potx" to ".pptx"
         *     ".potm" to ".pptm"
         * </p>
         * For example, a code converting  a .xlsm macro workbook to .xlsx would look as follows:
         * <p>
         *    <pre><code>
         *
         *     OPCPackage pkg = OPCPackage.open(new FileInputStream("macro-workbook.xlsm"));
         *     pkg.replaceContentType(
         *         "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
         *         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
         *
         *     FileOutputStream out = new FileOutputStream("workbook.xlsx");
         *     pkg.save(out);
         *     out.close();
         *
         *    </code></pre>
         * </p>
         *
         * @param oldContentType  the content type to be replaced
         * @param newContentType  the replacement
         * @return whether replacement was succesfull
         * @since POI-3.8
         */
        public bool ReplaceContentType(String oldContentType, String newContentType)
        {
            bool success = false;
            List<PackagePart> list = GetPartsByContentType(oldContentType);
            foreach (PackagePart packagePart in list)
            {
                if (packagePart.ContentType.Equals(oldContentType))
                {
                    PackagePartName partName = packagePart.PartName;
                    contentTypeManager.AddContentType(partName, newContentType);
                    success = true;
                }
            }
            return success;
        }
    }

}
