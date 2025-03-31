/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using System.Text.RegularExpressions;
using System.Collections.Generic;
using NPOI.XSSF.Model;
using NPOI.SS.UserModel;
using NPOI.Util;
using System.IO;
using System;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using System.Text; 
using Cysharp.Text;
using NPOI.SS.Util;
using NPOI.SS.Formula;
using NPOI.XSSF.UserModel.Helpers;
using NPOI.SS.Formula.UDF;
using NPOI.OpenXmlFormats;
using System.Collections;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.SS;
using System.Globalization;
using System.Linq;

namespace NPOI.XSSF.UserModel
{
    /**
     * High level representation of a SpreadsheetML workbook.  This is the first object most users
     * will construct whether they are Reading or writing a workbook.  It is also the
     * top level object for creating new sheets/etc.
     */
    public class XSSFWorkbook : POIXMLDocument, IWorkbook
    {
        private static readonly Regex COMMA_PATTERN = new Regex(",", RegexOptions.Compiled);

        /**
         * Width of one character of the default font in pixels. Same for Calibry and Arial.
         */
        public static float DEFAULT_CHARACTER_WIDTH = 7.0017f;

        /**
         * Excel silently tRuncates long sheet names to 31 chars.
         * This constant is used to ensure uniqueness in the first 31 chars
         */
        private static readonly int Max_SENSITIVE_SHEET_NAME_LEN = 31;
        /** Extended windows meta file */
        public static int PICTURE_TYPE_EMF = 2;

        /** Windows Meta File */
        public static int PICTURE_TYPE_WMF = 3;

        /** Mac PICT format */
        public static int PICTURE_TYPE_PICT = 4;

        /** JPEG format */
        public static int PICTURE_TYPE_JPEG = 5;

        /** PNG format */
        public static int PICTURE_TYPE_PNG = 6;

        /** Device independent bitmap */
        public static int PICTURE_TYPE_DIB = 7;
        /**
         * Images formats supported by XSSF but not by HSSF
         */
        public static int PICTURE_TYPE_GIF = 8;
        public static int PICTURE_TYPE_TIFF = 9;
        public static int PICTURE_TYPE_EPS = 10;
        public static int PICTURE_TYPE_BMP = 11;
        public static int PICTURE_TYPE_WPG = 12;
        public static int PICTURE_TYPE_JPG = 13; //Alias for JPEG to handle image/jpg content type
        /**
         * The underlying XML bean
         */
        private CT_Workbook workbook;

        /**
         * this holds the XSSFSheet objects attached to this workbook
         */
        private List<XSSFSheet> sheets;

        /**
         * this holds the XSSFName objects attached to this workbook, keyed by lower-case name
         */
        private Dictionary<String, List<XSSFName>> namedRangesByName;
        /**
         * this holds the XSSFName objects attached to this workbook
         */
        private List<XSSFName> namedRanges;

        /**
         * shared string table - a cache of strings in this workbook
         */
        private SharedStringsTable sharedStringSource;

        /**
         * A collection of shared objects used for styling content,
         * e.g. fonts, cell styles, colors, etc.
         */
        private StylesTable stylesSource;

        /**
         * The locator of user-defined functions.
         * By default includes functions from the Excel Analysis Toolpack
         */
        private readonly IndexedUDFFinder _udfFinder = new IndexedUDFFinder(UDFFinder.GetDefault());

        /**
         * TODO
         */
        private CalculationChain calcChain;

        /**
         * External Links, for referencing names or cells in other workbooks.
         */
        private List<ExternalLinksTable> externalLinks;

        /**
         * A collection of custom XML mappings
         */
        private MapInfo mapInfo;

        /**
         * Used to keep track of the data formatter so that all
         * CreateDataFormatter calls return the same one for a given
         * book.  This ensures that updates from one places is visible
         * someplace else.
         */
        private XSSFDataFormat formatter;

        /**
         * The policy to apply in the event of missing or
         *  blank cells when fetching from a row.
         * See {@link NPOI.ss.usermodel.Row.MissingCellPolicy}
         */
        private MissingCellPolicy _missingCellPolicy = MissingCellPolicy.RETURN_NULL_AND_BLANK;
        private bool cellFormulaValidation = true;
        /**
         * array of pictures for this workbook
         */
        private List<XSSFPictureData> pictures;

        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(XSSFWorkbook));

        /**
         * cached instance of XSSFCreationHelper for this workbook
         * @see {@link #getCreationHelper()}
         */
        private XSSFCreationHelper _creationHelper;

        /**
         * List of all pivot tables in workbook
         */
        private List<XSSFPivotTable> pivotTables;
        private List<CT_PivotCache> pivotCaches;
        /**
         * Create a new SpreadsheetML workbook.
         */
        public XSSFWorkbook()
            : this(XSSFWorkbookType.XLSX)
        {

        }

        /**
         * Create a new SpreadsheetML workbook.
         * @param workbookType The type of workbook to make (.xlsx or .xlsm).
         */
        public XSSFWorkbook(XSSFWorkbookType workbookType) :
                base(newPackage(workbookType))
        {

            OnWorkbookCreate();
        }

        /**
         * Constructs a XSSFWorkbook object given a OpenXML4J <code>Package</code> object,
         *  see <a href="http://poi.apache.org/oxml4j/">http://poi.apache.org/oxml4j/</a>.
         * 
         * Once you have finished working with the Workbook, you should close the package
         * by calling pkg.close, to avoid leaving file handles open.
         * 
         * Creating a XSSFWorkbook from a file-backed OPC Package has a lower memory
         *  footprint than an InputStream backed one.
         *
         * @param pkg the OpenXML4J <code>OPC Package</code> object.
         */
        public XSSFWorkbook(OPCPackage pkg)
            : base(pkg)
        {
            BeforeDocumentRead();

            //build a tree of POIXMLDocumentParts, this workbook being the root
            Load(XSSFFactory.GetInstance());

            // some broken Workbooks miss this...
            if (!workbook.IsSetBookViews())
            {
                CT_BookViews bvs = workbook.AddNewBookViews();
                CT_BookView bv = bvs.AddNewWorkbookView();
                bv.activeTab = (0);
            }
        }
        /**
         * Constructs a XSSFWorkbook object, by buffering the whole stream into memory
         *  and then opening an {@link OPCPackage} object for it.
         * 
         * Using an {@link InputStream} requires more memory than using a File, so
         *  if a {@link File} is available then you should instead do something like
         *   <pre><code>
         *       OPCPackage pkg = OPCPackage.open(path);
         *       XSSFWorkbook wb = new XSSFWorkbook(pkg);
         *       // work with the wb object
         *       ......
         *       pkg.close(); // gracefully closes the underlying zip file
         *   </code></pre>     
         */
        public XSSFWorkbook(Stream fileStream, bool readOnly = false)
            : base(PackageHelper.Open(fileStream, readOnly))
        {
            BeforeDocumentRead();

            //build a tree of POIXMLDocumentParts, this workbook being the root
            Load(XSSFFactory.GetInstance());

            // some broken Workbooks miss this...
            if (!workbook.IsSetBookViews())
            {
                CT_BookViews bvs = workbook.AddNewBookViews();
                CT_BookView bv = bvs.AddNewWorkbookView();
                bv.activeTab = (0);
            }
        }

        /**
         * Constructs a XSSFWorkbook object from a given file.
         * 
         * <p>Once you have finished working with the Workbook, you should close 
         * the package by calling  {@link #close()}, to avoid leaving file 
         * handles open.
         * 
         * <p>Opening a XSSFWorkbook from a file has a lower memory footprint 
         *  than opening from an InputStream
         *  
         * @param file   the file to open
         */
        public XSSFWorkbook(FileInfo file, bool readOnly = false)
            : this(OPCPackage.Open(file, readOnly? PackageAccess.READ: PackageAccess.READ_WRITE))
        {

        }

        /**
         * Constructs a XSSFWorkbook object given a file name.
         *
         * <p>
         *  This constructor is deprecated since POI-3.8 because it does not close
         *  the underlying .zip file stream. In short, there are two ways to open a OPC package:
         * </p>
         * <ol>
         *     <li>
         *      from file which leads to invoking java.util.zip.ZipFile(File file)
         *      deep in POI internals.
         *     </li>
         *     <li>
         *     from input stream in which case we first read everything into memory and
         *     then pass the data to ZipInputStream.
         *     </li>
         * <ol>
         * <p>    
         *     It should be noted, that (2) uses quite a bit more memory than (1), which
         *      doesn't need to hold the whole zip file in memory, and can take advantage
         *      of native methods.
         * </p>
         * <p>
         *   To construct a workbook from file use the
         *   {@link #XSSFWorkbook(org.apache.poi.openxml4j.opc.OPCPackage)}  constructor:
         *   <pre><code>
         *       OPCPackage pkg = OPCPackage.open(path);
         *       XSSFWorkbook wb = new XSSFWorkbook(pkg);
         *       // work with the wb object
         *       ......
         *       pkg.close(); // gracefully closes the underlying zip file
         *   </code></pre>     
         * </p>
         * 
         * @param      path   the file name.
         */
        public XSSFWorkbook(String path, bool readOnly = false)
            : this(OpenPackage(path, readOnly))
        {

        }

        protected void BeforeDocumentRead()
        {
            // Ensure it isn't a XLSB file, which we don't support
            if (CorePart.ContentType.Equals(XSSFRelation.XLSB_BINARY_WORKBOOK.ContentType))
            {
                throw new XLSBUnsupportedException();
            }

            // Create arrays for parts attached to the workbook itself
            pivotTables = new List<XSSFPivotTable>();
            pivotCaches = new List<CT_PivotCache>();
        }

        WorkbookDocument doc = null;
        internal override void OnDocumentRead()
        {
            try
            {
                XmlDocument xmldoc = ConvertStreamToXml(GetPackagePart().GetInputStream());
                doc = WorkbookDocument.Parse(xmldoc, NamespaceManager);
                this.workbook = doc.Workbook;

                ThemesTable theme = null;
                Dictionary<String, XSSFSheet> shIdMap = new Dictionary<String, XSSFSheet>();
                Dictionary<String, ExternalLinksTable> elIdMap = new Dictionary<String, ExternalLinksTable>();
                foreach (RelationPart rp in RelationParts)
                {
                    POIXMLDocumentPart p = rp.DocumentPart;
                    if (p is SharedStringsTable table) sharedStringSource = table;
                    else if (p is StylesTable stylesTable) stylesSource = stylesTable;
                    else if (p is ThemesTable themesTable) theme = themesTable;
                    else if (p is CalculationChain chain) calcChain = chain;
                    else if (p is MapInfo info) mapInfo = info;
                    else if (p is XSSFSheet sheet)
                    {
                        shIdMap[rp.Relationship.Id] = sheet;
                    }
                    else if (p is ExternalLinksTable linksTable)
                    {
                        elIdMap[rp.Relationship.Id] = linksTable;
                    }
                }

                bool packageReadOnly = (Package.GetPackageAccess() == PackageAccess.READ);

                if (stylesSource == null)
                {
                    // Create Styles if it is missing
                    if (packageReadOnly)
                    {
                        stylesSource = new StylesTable();
                    }
                    else
                    {
                        stylesSource = (StylesTable)CreateRelationship(XSSFRelation.STYLES, XSSFFactory.GetInstance());
                    }
                }
                stylesSource.SetWorkbook(this);
                stylesSource.SetTheme(theme);

                if (sharedStringSource == null)
                {
                    //Create SST if it is missing
                    if (packageReadOnly)
                    {
                        sharedStringSource = new SharedStringsTable();
                    }
                    else
                    {
                        sharedStringSource = (SharedStringsTable)CreateRelationship(XSSFRelation.SHARED_STRINGS, XSSFFactory.GetInstance());
                    }
                }

                // Load individual sheets. The order of sheets is defined by the order
                //  of CTSheet elements in the workbook
                sheets = new List<XSSFSheet>(shIdMap.Count);
                foreach (CT_Sheet ctSheet in this.workbook.sheets.sheet)
                {
                    ParseSheet(shIdMap, ctSheet);

                }
                // Load the external links tables. Their order is defined by the order 
                //  of CTExternalReference elements in the workbook
                externalLinks = new List<ExternalLinksTable>(elIdMap.Count);
                if (this.workbook.IsSetExternalReferences())
                {
                    foreach (CT_ExternalReference er in this.workbook.externalReferences.externalReference)
                    {
                        ExternalLinksTable el = null;
                        if (elIdMap.TryGetValue(er.id, out ExternalLinksTable value))
                            el = value;
                        if (el == null)
                        {
                            logger.Log(POILogger.WARN, "ExternalLinksTable with r:id " + er.id + " was defined, but didn't exist in package, skipping");
                            continue;
                        }
                        externalLinks.Add(el);
                    }
                }
                // Process the named ranges
                ReprocessNamedRanges();
            }
            catch (XmlException e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * Not normally to be called externally, but possibly to be overridden to avoid
         * the DOM based parse of large sheets (see examples).
         */
        private void ParseSheet(Dictionary<String, XSSFSheet> shIdMap, CT_Sheet ctSheet)
        {
            XSSFSheet sh = null;
            if (shIdMap.TryGetValue(ctSheet.id, out XSSFSheet value))
                sh = value;
            if (sh == null)
            {
                logger.Log(POILogger.WARN, "Sheet with name " + ctSheet.name + " and r:id " + ctSheet.id + " was defined, but didn't exist in package, skipping");
                return;
            }
            sh.sheet = ctSheet;
            sh.OnDocumentRead();
            sheets.Add(sh);
        }
        /**
         * Create a new CT_Workbook with all values Set to default
         */
        private void OnWorkbookCreate()
        {

            doc = new WorkbookDocument();
            workbook = doc.Workbook;
            // don't EVER use the 1904 date system
            CT_WorkbookPr workbookPr = workbook.AddNewWorkbookPr();
            workbookPr.date1904 = (false);

            CT_BookViews bvs = workbook.AddNewBookViews();
            CT_BookView bv = bvs.AddNewWorkbookView();
            bv.activeTab = 0;
            workbook.AddNewSheets();

            ExtendedProperties expProps = GetProperties().ExtendedProperties;
            CT_ExtendedProperties ctExtendedProp = expProps.GetUnderlyingProperties();
            ctExtendedProp.Application = DOCUMENT_CREATOR;
            ctExtendedProp.DocSecurity = 0;
            ctExtendedProp.DocSecuritySpecified = true;
            ctExtendedProp.ScaleCrop = false;
            ctExtendedProp.ScaleCropSpecified = true;
            ctExtendedProp.LinksUpToDate = false;
            ctExtendedProp.LinksUpToDateSpecified = true;
            ctExtendedProp.HyperlinksChanged = false;
            ctExtendedProp.HyperlinksChangedSpecified = true;
            ctExtendedProp.SharedDoc = false;
            ctExtendedProp.SharedDocSpecified = true;

            sharedStringSource = (SharedStringsTable)CreateRelationship(XSSFRelation.SHARED_STRINGS, XSSFFactory.GetInstance());
            stylesSource = (StylesTable)CreateRelationship(XSSFRelation.STYLES, XSSFFactory.GetInstance());
            stylesSource.SetWorkbook(this);

            namedRanges = new List<XSSFName>();
            namedRangesByName = new Dictionary<string, List<XSSFName>>();
            sheets = new List<XSSFSheet>();

            pivotTables = new List<XSSFPivotTable>();
        }

        /**
         * Create a new SpreadsheetML namespace and Setup the default minimal content
         */
        protected static OPCPackage newPackage(XSSFWorkbookType workbookType)
        {
            try
            {
                OPCPackage pkg = OPCPackage.Create(new MemoryStream());
                // Main part
                PackagePartName corePartName = PackagingUriHelper.CreatePartName(XSSFRelation.WORKBOOK.DefaultFileName);
                // Create main part relationship
                pkg.AddRelationship(corePartName, TargetMode.Internal, PackageRelationshipTypes.CORE_DOCUMENT);
                // Create main document part
                pkg.CreatePart(corePartName, workbookType.ContentType);

                pkg.GetPackageProperties().SetCreatorProperty(DOCUMENT_CREATOR);

                return pkg;
            }
            catch (Exception e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * Return the underlying XML bean
         *
         * @return the underlying CT_Workbook bean
         */

        public CT_Workbook GetCTWorkbook()
        {
            return this.workbook;
        }

        /**
         * Adds a picture to the workbook.
         *
         * @param pictureData       The bytes of the picture
         * @param format            The format of the picture.
         *
         * @return the index to this picture (0 based), the Added picture can be obtained from {@link #getAllPictures()} .
         * @see Workbook#PICTURE_TYPE_EMF
         * @see Workbook#PICTURE_TYPE_WMF
         * @see Workbook#PICTURE_TYPE_PICT
         * @see Workbook#PICTURE_TYPE_JPEG
         * @see Workbook#PICTURE_TYPE_PNG
         * @see Workbook#PICTURE_TYPE_DIB
         * @see #getAllPictures()
         */
        public int AddPicture(byte[] pictureData, int format)
        {
            return this.AddPicture(pictureData, (PictureType)format);
        }

        /**
         * Adds a picture to the workbook.
         *
         * @param is                The sream to read image from
         * @param format            The format of the picture.
         *
         * @return the index to this picture (0 based), the Added picture can be obtained from {@link #getAllPictures()} .
         * @see Workbook#PICTURE_TYPE_EMF
         * @see Workbook#PICTURE_TYPE_WMF
         * @see Workbook#PICTURE_TYPE_PICT
         * @see Workbook#PICTURE_TYPE_JPEG
         * @see Workbook#PICTURE_TYPE_PNG
         * @see Workbook#PICTURE_TYPE_DIB
         * @see #getAllPictures()
         */
        public int AddPicture(Stream picStream, int format)
        {
            int imageNumber = GetAllPictures().Count + 1;
            XSSFPictureData img = (XSSFPictureData)CreateRelationship(XSSFPictureData.RELATIONS[format], XSSFFactory.GetInstance(), imageNumber, true).DocumentPart;
            Stream out1 = img.GetPackagePart().GetOutputStream();
            IOUtils.Copy(picStream, out1);
            out1.Close();
            pictures.Add(img);
            return imageNumber - 1;
        }

        /**
         * Create an XSSFSheet from an existing sheet in the XSSFWorkbook.
         *  The Cloned sheet is a deep copy of the original.
         *  
         * @param sheetNum The index of the sheet to clone
         * @return XSSFSheet representing the Cloned sheet.
         * @throws ArgumentException if the sheet index in invalid
         * @throws POIXMLException if there were errors when cloning
         */
        public ISheet CloneSheet(int sheetNum)
        {
            return CloneSheet(sheetNum, null);
        }

        /**
         * Create an XSSFSheet from an existing sheet in the XSSFWorkbook.
         *  The cloned sheet is a deep copy of the original but with a new given
         *  name.
         *
         * @param sheetNum The index of the sheet to clone
         * @param newName The name to set for the newly created sheet
         * @return XSSFSheet representing the cloned sheet.
         * @throws IllegalArgumentException if the sheet index or the sheet
         *         name is invalid
         * @throws POIXMLException if there were errors when cloning
         */
        public ISheet CloneSheet(int sheetNum, String newName)
        {
            ValidateSheetIndex(sheetNum);

            XSSFSheet srcSheet = sheets[sheetNum];
            //return srcSheet.CopySheet(srcSheet.SheetName);
            if (newName == null)
            {
                String srcName = srcSheet.SheetName;
                newName = GetUniqueSheetName(srcName);
            }
            else
            {
                ValidateSheetName(newName);
                WorkbookUtil.ValidateSheetName(newName);
            }

            XSSFSheet clonedSheet = CreateSheet(newName) as XSSFSheet;

            // copy sheet's relations
            IList<RelationPart> rels = srcSheet.RelationParts;
            // if the sheet being cloned has a drawing then rememebr it and re-create it too
            XSSFDrawing dg = null;
            foreach (RelationPart rp in rels)
            {
                POIXMLDocumentPart r = rp.DocumentPart;
                // do not copy the drawing relationship, it will be re-created
                if (r is XSSFDrawing drawing)
                {
                    dg = drawing;
                    continue;
                }

                AddRelation(rp, clonedSheet);
            }

            try
            {
                foreach (PackageRelationship pr in srcSheet.GetPackagePart().Relationships)
                {
                    if (pr.TargetMode == TargetMode.External)
                    {
                        clonedSheet.GetPackagePart().AddExternalRelationship
                            (pr.TargetUri.OriginalString, pr.RelationshipType, pr.Id);
                    }
                }
            }
            catch (InvalidFormatException e)
            {
                throw new POIXMLException("Failed to clone sheet", e);
            }

            try
            {
                using (MemoryStream ms = RecyclableMemory.GetStream())
                {
                    srcSheet.Write(ms, true);
                    ms.Position = 0;
                    clonedSheet.Read(ms);
                }
            }
            catch (IOException e)
            {
                throw new POIXMLException("Failed to clone sheet", e);
            }

            CT_Worksheet ct = clonedSheet.GetCTWorksheet();
            if (ct.IsSetLegacyDrawing())
            {
                //logger.log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
                ct.UnsetLegacyDrawing();
            }
            if (ct.IsSetPageSetup())
            {
                //logger.log(POILogger.WARN, "Cloning sheets with page setup is not yet supported.");
                ct.UnsetPageSetup();
            }

            if (srcSheet.RepeatingRows != null)
                clonedSheet.RepeatingRows = srcSheet.RepeatingRows;

            if (srcSheet.RepeatingColumns != null)
                clonedSheet.RepeatingColumns = srcSheet.RepeatingColumns;

            clonedSheet.IsSelected = (false);

            // clone the sheet drawing alongs with its relationships
            if (dg != null)
            {
                if (ct.IsSetDrawing())
                {
                    // unset the existing reference to the drawing,
                    // so that subsequent call of clonedSheet.createDrawingPatriarch() will create a new one
                    ct.UnsetDrawing();
                }
                XSSFDrawing clonedDg = clonedSheet.CreateDrawingPatriarch() as XSSFDrawing;
                // copy drawing contents
                clonedDg.GetCTDrawing().Set(dg.GetCTDrawing());

                clonedDg = clonedSheet.CreateDrawingPatriarch() as XSSFDrawing;

                // Clone drawing relations
                IList<RelationPart> srcRels = (srcSheet.CreateDrawingPatriarch() as XSSFDrawing).RelationParts;
                foreach (RelationPart rp in srcRels)
                {
                    AddRelation(rp, clonedDg);
                }
            }
            return clonedSheet;
        }

        /**
         * @since 3.14-Beta1
         */
        private static void AddRelation(RelationPart rp, POIXMLDocumentPart target)
        {
            PackageRelationship rel = rp.Relationship;
            if (rel.TargetMode == TargetMode.External)
            {
                target.GetPackagePart().AddRelationship(
                    rel.TargetUri, rel.TargetMode.Value, rel.RelationshipType, rel.Id);
            }
            else
            {
                XSSFRelation xssfRel = XSSFRelation.GetInstance(rel.RelationshipType);
                if (xssfRel == null)
                {
                    // Don't copy all relations blindly, but only the ones we know about
                    throw new POIXMLException("Can't clone sheet - unknown relation type found: " + rel.RelationshipType);
                }
                target.AddRelation(rel.Id, xssfRel, rp.DocumentPart);
            }
        }

        /**
         * Generate a valid sheet name based on the existing one. Used when cloning sheets.
         *
         * @param srcName the original sheet name to
         * @return clone sheet name
         */
        private String GetUniqueSheetName(String srcName)
        {
            int uniqueIndex = 2;
            String baseName = srcName;
            int bracketPos = srcName.LastIndexOf('(');
            if (bracketPos > 0 && srcName.EndsWith(")"))
            {
                String suffix = srcName.Substring(bracketPos + 1, srcName.Length - ")".Length - bracketPos - 1);
                try
                {
                    uniqueIndex = int.Parse(suffix.Trim());
                    uniqueIndex++;
                    baseName = srcName.Substring(0, bracketPos).Trim();
                }
                catch (FormatException)
                {
                    // contents of brackets not numeric
                }
            }
            while (true)
            {
                // Try and find the next sheet name that is unique
                String index = (uniqueIndex++).ToString();
                String name;
                if (baseName.Length + index.Length + 2 < 31)
                {
                    name = baseName + " (" + index + ")";
                }
                else
                {
                    name = baseName.Substring(0, 31 - index.Length - 2) + "(" + index + ")";
                }

                //If the sheet name is unique, then set it otherwise move on to the next number.
                if (GetSheetIndex(name) == -1)
                {
                    return name;
                }
            }
        }
        /// <summary>
        /// Create a new XSSFCellStyle and add it to the workbook's style table
        /// </summary>
        /// <returns>the new XSSFCellStyle object</returns>
        public ICellStyle CreateCellStyle()
        {
            return stylesSource.CreateCellStyle();
        }

        /// <summary>
        /// Returns the workbook's data format table (a factory for creating data format strings).
        /// </summary>
        /// <returns>the XSSFDataFormat object</returns>
        public IDataFormat CreateDataFormat()
        {
            if (formatter == null)
                formatter = new XSSFDataFormat(stylesSource);
            return formatter;
        }
        /// <summary>
        /// Create a new Font and add it to the workbook's font table
        /// </summary>
        /// <returns></returns>
        public IFont CreateFont()
        {
            XSSFFont font = new XSSFFont();
            font.RegisterTo(stylesSource);
            return font;
        }

        public IName CreateName()
        {
            CT_DefinedName ctName = new CT_DefinedName();
            ctName.name = ("");
            return CreateAndStoreName(ctName);
        }

        private void PutValuesMapping(string key, XSSFName name)
        {
            if (namedRangesByName.ContainsKey(key))
                namedRangesByName[key].Add(name);
            else
                namedRangesByName.Add(key, new List<XSSFName>() { name });
        }

        private XSSFName CreateAndStoreName(CT_DefinedName ctName)
        {
            XSSFName name = new XSSFName(ctName, this);
            namedRanges.Add(name);
            PutValuesMapping(ctName.name.ToLower(), name);
            return name;
        }

        /**
         * Create an XSSFSheet for this workbook, Adds it to the sheets and returns
         * the high level representation.  Use this to create new sheets.
         *
         * @return XSSFSheet representing the new sheet.
         */
        public ISheet CreateSheet()
        {
            String sheetname = "Sheet" + (sheets.Count);
            int idx = 0;
            while (GetSheet(sheetname) != null)
            {
                sheetname = "Sheet" + idx;
                idx++;
            }
            return CreateSheet(sheetname);
        }

        /**
         * Create a new sheet for this Workbook and return the high level representation.
         * Use this to create new sheets.
         *
         * <p>
         *     Note that Excel allows sheet names up to 31 chars in length but other applications
         *     (such as OpenOffice) allow more. Some versions of Excel crash with names longer than 31 chars,
         *     others - tRuncate such names to 31 character.
         * </p>
         * <p>
         *     POI's SpreadsheetAPI silently tRuncates the input argument to 31 characters.
         *     Example:
         *
         *     <pre><code>
         *     Sheet sheet = workbook.CreateSheet("My very long sheet name which is longer than 31 chars"); // will be tRuncated
         *     assert 31 == sheet.SheetName.Length;
         *     assert "My very long sheet name which i" == sheet.SheetName;
         *     </code></pre>
         * </p>
         *
         * Except the 31-character constraint, Excel applies some other rules:
         * <p>
         * Sheet name MUST be unique in the workbook and MUST NOT contain the any of the following characters:
         * <ul>
         * <li> 0x0000 </li>
         * <li> 0x0003 </li>
         * <li> colon (:) </li>
         * <li> backslash (\) </li>
         * <li> asterisk (*) </li>
         * <li> question mark (?) </li>
         * <li> forward slash (/) </li>
         * <li> opening square bracket ([) </li>
         * <li> closing square bracket (]) </li>
         * </ul>
         * The string MUST NOT begin or end with the single quote (') character.
         * </p>
         *
         * <p>
         * See {@link org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String nameProposal)}
         *      for a safe way to create valid names
         * </p>
         * @param sheetname  sheetname to set for the sheet.
         * @return Sheet representing the new sheet.
         * @throws IllegalArgumentException if the name is null or invalid
         *  or workbook already contains a sheet with this name
         * @see org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String nameProposal)
         */
        public ISheet CreateSheet(String sheetname)
        {
            if (sheetname == null)
            {
                throw new ArgumentException("sheetName must not be null");
            }
            ValidateSheetName(sheetname);

            WorkbookUtil.ValidateSheetName(sheetname);

            CT_Sheet sheet = AddSheet(sheetname);

            int sheetNumber = 1;
            //TODO: this is extra somehow
            foreach (XSSFSheet sh in sheets) sheetNumber = (int)Math.Max(sh.sheet.sheetId + 1, sheetNumber);

            outerloop:
            while (true)
            {
                foreach (XSSFSheet sh in sheets)
                {
                    sheetNumber = (int)Math.Max(sh.sheet.sheetId + 1, sheetNumber);
                }

                // Bug 57165: We also need to check that the resulting file name is not already taken
                // this can happen when moving/cloning sheets
                String sheetName = XSSFRelation.WORKSHEET.GetFileName(sheetNumber);
                foreach (POIXMLDocumentPart relation in GetRelations())
                {
                    if (relation.GetPackagePart() != null &&
                            sheetName.Equals(relation.GetPackagePart().PartName.Name))
                    {
                        // name is taken => try next one
                        sheetNumber++;
                        goto outerloop;
                    }
                }

                // no duplicate found => use this one
                break;
            }

            RelationPart rp = CreateRelationship(XSSFRelation.WORKSHEET, XSSFFactory.GetInstance(), sheetNumber, false);
            XSSFSheet wrapper = rp.DocumentPart as XSSFSheet;
            wrapper.sheet = sheet;
            sheet.id = (rp.Relationship.Id);
            sheet.sheetId = (uint)sheetNumber;
            if (sheets.Count == 0) wrapper.IsSelected = (true);
            sheets.Add(wrapper);
            return wrapper;
        }
        private void ValidateSheetName(String sheetName)
        {
            if (ContainsSheet(sheetName, sheets.Count))
                throw new ArgumentException(string.Format("The workbook already contains a sheet named '{0}'", sheetName));
        }
        protected XSSFDialogsheet CreateDialogsheet(String sheetname, CT_Dialogsheet dialogsheet)
        {
            XSSFSheet sheet = CreateSheet(sheetname) as XSSFSheet;
            String sheetRelId = GetRelationId(sheet);
            PackageRelationship pr = GetPackagePart().GetRelationship(sheetRelId);
            return new XSSFDialogsheet(sheet, pr);
        }

        private CT_Sheet AddSheet(String sheetname)
        {
            CT_Sheet sheet = workbook.sheets.AddNewSheet();
            sheet.name = (sheetname);
            return sheet;
        }

        /**
         * Finds a font that matches the one with the supplied attributes
         */
        [Obsolete("deprecated POI 3.15. Use {@link #findFont(boolean, short, short, String, boolean, boolean, short, byte)} instead.")]
        public IFont FindFont(short boldWeight, short color, short fontHeight, String name, bool italic, bool strikeout, FontSuperScript typeOffset, FontUnderlineType underline)
        {
            return stylesSource.FindFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
        }
        /**
         * Finds a font that matches the one with the supplied attributes
         */
        public IFont FindFont(bool bold, short color, short fontHeight, String name, bool italic, bool strikeout, FontSuperScript typeOffset, FontUnderlineType underline)
        {
            return stylesSource.FindFont(bold, color, fontHeight, name, italic, strikeout, typeOffset, underline);
        }
        /**
         * Convenience method to Get the active sheet.  The active sheet is is the sheet
         * which is currently displayed when the workbook is viewed in Excel.
         * 'Selected' sheet(s) is a distinct concept.
         */
        public int ActiveSheetIndex
        {
            get
            {
                //activeTab (Active Sheet Index) Specifies an unsignedInt
                //that Contains the index to the active sheet in this book view.
                return (int)workbook.bookViews.GetWorkbookViewArray(0).activeTab;
            }
        }

        /**
         * Gets all pictures from the Workbook.
         *
         * @return the list of pictures (a list of {@link XSSFPictureData} objects.)
         * @see #AddPicture(byte[], int)
         */
        public IList GetAllPictures()
        {
            if (pictures == null)
            {
                List<PackagePart> mediaParts = Package.GetPartsByName(new Regex("/xl/media/.*?"));
                pictures = new List<XSSFPictureData>(mediaParts.Count);
                foreach (PackagePart part in mediaParts)
                {
                    pictures.Add(new XSSFPictureData(part));
                }
            }
            return pictures;
        }

        /**
         * Get the cell style object at the given index
         *
         * @param idx  index within the set of styles
         * @return XSSFCellStyle object at the index
         */
        public ICellStyle GetCellStyleAt(int idx)
        {
            return stylesSource.GetStyleAt(idx);
        }

        /**
         * Get the font at the given index number
         *
         * @param idx  index number
         * @return XSSFFont at the index
         */
        public IFont GetFontAt(short idx)
        {
            return stylesSource.GetFontAt(idx);
        }

        /// <summary>
        /// Get the first named range with the given name.
        /// Note: names of named ranges are not unique as they are scoped by sheet.
        /// {@link #getNames(String name)} returns all named ranges with the given name.
        /// </summary>
        /// <param name="name">named range name</param>
        /// <returns>return XSSFName with the given name. <code>null</code> is returned no named range could be found.</returns>
        public IName GetName(String name)
        {
            IList<IName> list = GetNames(name);
            if (list.Count == 0)
            {
                return null;
            }
            return list[0];
        }
        /// <summary>
        /// Get the named ranges with the given name.
        /// <i>Note:</i>Excel named ranges are case-insensitive and
        /// this method performs a case-insensitive search.
        /// </summary>
        /// <param name="name">named range name</param>
        /// <returns>return list of XSSFNames with the given name. An empty list if no named ranges could be found</returns>
        public IList<IName> GetNames(String name)
        {
            var ret = new List<IName>();
            if (namedRangesByName.ContainsKey(name.ToLower()))
            {
                ret.AddRange(namedRangesByName[name.ToLower()]);
            }
            return ret.AsReadOnly();
            //return Collections.unmodifiableList(namedRangesByName.get(name.toLowerCase(Locale.ENGLISH)));
        }
        [Obsolete("deprecated 3.16. New projects should avoid accessing named ranges by index.")]
        public IName GetNameAt(int nameIndex)
        {
            int nNames = namedRanges.Count;
            if (nNames < 1)
            {
                throw new InvalidOperationException("There are no defined names in this workbook");
            }
            if (nameIndex < 0 || nameIndex > nNames)
            {
                throw new ArgumentException("Specified name index " + nameIndex
                        + " is outside the allowable range (0.." + (nNames - 1) + ").");
            }
            return namedRanges[nameIndex];
        }

        /// <summary>
        /// Get a list of all the named ranges in the workbook.
        /// </summary>
        /// <returns>return list of XSSFNames in the workbook</returns>
        public IList<IName> GetAllNames()
        {
            var ret = new List<IName>();
            ret.AddRange(namedRanges);
            return ret.AsReadOnly();
        }
        /**
         * Gets the named range index by his name
         * <i>Note:</i>Excel named ranges are case-insensitive and
         * this method performs a case-insensitive search.
         *
         * @param name named range name
         * @return named range index
         */
        [Obsolete("deprecated 3.16. New projects should avoid accessing named ranges by index. Use {@link #getName(String)} instead.")]
        public int GetNameIndex(String name)
        {
            XSSFName nm = GetName(name) as XSSFName;
            if (nm != null)
            {
                return namedRanges.IndexOf(nm);
            }
            return -1;
        }

        /**
         * Get the number of styles the workbook Contains
         *
         * @return count of cell styles
         */
        public int NumCellStyles
        {
            get
            {
                return stylesSource.NumCellStyles;
            }
        }

        /**
         * Get the number of fonts in the this workbook
         *
         * @return number of fonts
         */
        public short NumberOfFonts
        {
            get
            {
                return (short)stylesSource.GetFonts().Count;
            }
        }

        /**
         * Get the number of named ranges in the this workbook
         *
         * @return number of named ranges
         */
        public int NumberOfNames
        {
            get
            {
                return namedRanges.Count;
            }
        }

        /**
         * Get the number of worksheets in the this workbook
         *
         * @return number of worksheets
         */
        public int NumberOfSheets
        {
            get
            {
                return sheets.Count;
            }
        }

        /**
         * Retrieves the reference for the printarea of the specified sheet, the sheet name is Appended to the reference even if it was not specified.
         * @param sheetIndex Zero-based sheet index (0 Represents the first sheet to keep consistent with java)
         * @return String Null if no print area has been defined
         */
        public String GetPrintArea(int sheetIndex)
        {
            XSSFName name = GetBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
            if (name == null) return null;
            //Adding one here because 0 indicates a global named region; doesnt make sense for print areas
            return name.RefersToFormula;

        }

        /**
         * Get sheet with the given name (case insensitive match)
         *
         * @param name of the sheet
         * @return XSSFSheet with the name provided or <code>null</code> if it does not exist
         */
        public ISheet GetSheet(String name)
        {
            foreach (XSSFSheet sheet in sheets)
            {
                if (name.Equals(sheet.SheetName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return sheet;
                }
            }
            return null;
        }

        /**
         * Get the XSSFSheet object at the given index.
         *
         * @param index of the sheet number (0-based physical & logical)
         * @return XSSFSheet at the provided index
         * @throws ArgumentException if the index is out of range (index
         *            &lt; 0 || index &gt;= NumberOfSheets).
         */
        public ISheet GetSheetAt(int index)
        {
            ValidateSheetIndex(index);
            return sheets[index];
        }

        /// <summary>
        /// Returns the index of the sheet by his name (case insensitive match)
        /// </summary>
        /// <param name="name">the sheet name</param>
        /// <returns>index of the sheet (0 based) or -1 if not found</returns>
        public int GetSheetIndex(String name)
        {
            int idx = 0;
            foreach (XSSFSheet sh in sheets)
            {
                if (name.Equals(sh.SheetName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return idx;
                }
                idx++;
            }
            return -1;
        }

        /**
         * Returns the index of the given sheet
         *
         * @param sheet the sheet to look up
         * @return index of the sheet (0 based). <tt>-1</tt> if not found
         */
        public int GetSheetIndex(ISheet sheet)
        {
            int idx = 0;
            foreach (XSSFSheet sh in sheets)
            {
                if (sh == sheet) return idx;
                idx++;
            }
            return -1;
        }

        /**
         * Get the sheet name
         *
         * @param sheetIx Number
         * @return Sheet name
         */
        public String GetSheetName(int sheetIx)
        {
            ValidateSheetIndex(sheetIx);
            return sheets[sheetIx].SheetName;
        }

        /**
         * Allows foreach loops:
         * <pre><code>
         * XSSFWorkbook wb = new XSSFWorkbook(package);
         * for(XSSFSheet sheet : wb){
         *
         * }
         * </code></pre>
         */
        public IEnumerator<ISheet> GetEnumerator()
        {
            return sheets.GetEnumerator();
        }
        /**
         * Are we a normal workbook (.xlsx), or a
         *  macro enabled workbook (.xlsm)?
         */
        public bool IsMacroEnabled()
        {
            return GetPackagePart().ContentType.Equals(XSSFRelation.MACROS_WORKBOOK.ContentType);
        }
        [Obsolete("deprecated 3.16. New projects should use {@link #removeName(Name)}.")]
        public void RemoveName(int nameIndex)
        {
            RemoveName(GetNameAt(nameIndex));
            //namedRanges.RemoveAt(nameIndex);
        }
        public void RemoveName(String name)
        {
            List<XSSFName> names = namedRangesByName[name.ToLower()];
            if (names.Count == 0)
            {
                throw new ArgumentException("Named range was not found: " + name);
            }
            RemoveName(names[0]);
        }

        private bool RemoveMapping(string key, XSSFName item)
        {
            if (namedRangesByName.TryGetValue(key, out List<XSSFName> values))
            {
                return values.Remove(item);
            }
            return false;
        }
        /**
         * As {@link #removeName(String)} is not necessarily unique 
         * (name + sheet index is unique), this method is more accurate.
         * 
         * @param name the name to remove.
         */
        public void RemoveName(IName name)
        {
            if (!RemoveMapping(name.NameName.ToLower(), name as XSSFName))
            {
                throw new ArgumentException("Name was not found: " + name);
            }
            if (!namedRanges.Remove((XSSFName)name))
            {
                throw new ArgumentException("Name was not found: " + name);
            }
        }
        internal void UpdateName(XSSFName name, String oldName)
        {
            if (!RemoveMapping(oldName.ToLower(), name))
            {
                throw new ArgumentException("Name was not found: " + name);
            }

            PutValuesMapping(name.NameName.ToLower(), name);
        }
        /**
         * Delete the printarea for the sheet specified
         *
         * @param sheetIndex 0-based sheet index (0 = First Sheet)
         */
        public void RemovePrintArea(int sheetIndex)
        {
            XSSFName name = GetBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
            if (name != null)
            {
                RemoveName(name);
            }
        }

        /**
         * Removes sheet at the given index.<p/>
         *
         * Care must be taken if the Removed sheet is the currently active or only selected sheet in
         * the workbook. There are a few situations when Excel must have a selection and/or active
         * sheet. (For example when printing - see Bug 40414).<br/>
         *
         * This method Makes sure that if the Removed sheet was active, another sheet will become
         * active in its place.  Furthermore, if the Removed sheet was the only selected sheet, another
         * sheet will become selected.  The newly active/selected sheet will have the same index, or
         * one less if the Removed sheet was the last in the workbook.
         *
         * @param index of the sheet  (0-based)
         */
        public void RemoveSheetAt(int index)
        {
            ValidateSheetIndex(index);
            OnSheetDelete(index);

            XSSFSheet sheet = (XSSFSheet)GetSheetAt(index);
            RemoveRelation(sheet);
            sheets.RemoveAt(index);

            // only Set new sheet if there are still some left
            if (sheets.Count == 0)
            {
                return;
            }

            // the index of the closest remaining sheet to the one just deleted
            int newSheetIndex = index;
            if (newSheetIndex >= sheets.Count)
            {
                newSheetIndex = sheets.Count - 1;
            }

            // adjust active sheet
            int active = ActiveSheetIndex;
            if (active == index)
            {
                // Removed sheet was the active one, reset active sheet if there is still one left now
                SetActiveSheet(newSheetIndex);
            }
            else if (active > index)
            {
                // Removed sheet was below the active one => active is one less now
                SetActiveSheet(active - 1);
            }

        }

        /**
         * Gracefully remove references to the sheet being deleted
         *
         * @param index the 0-based index of the sheet to delete
         */
        private void OnSheetDelete(int index)
        {
            //delete the CT_Sheet reference from workbook.xml
            workbook.sheets.RemoveSheet(index);

            //calculation chain is auxiliary, remove it as it may contain orphan references to deleted cells
            if (calcChain != null)
            {
                RemoveRelation(calcChain);
                calcChain = null;
            }

            List<XSSFName> toRemove = new List<XSSFName>();
            foreach (XSSFName nm in namedRanges)
            {

                CT_DefinedName ct = nm.GetCTName();
                if (!ct.IsSetLocalSheetId()) continue;
                if (ct.localSheetId == index)
                {
                    toRemove.Add(nm);
                }
                else if (ct.localSheetId > index)
                {
                    // Bump down by one, so still points at the same sheet
                    ct.localSheetId = (ct.localSheetId - 1);
                    ct.localSheetIdSpecified = true;
                }
            }
            foreach (XSSFName nm in toRemove)
            {
                RemoveName(nm);
            }
        }

        /**
         * Retrieves the current policy on what to do when
         *  Getting missing or blank cells from a row.
         * The default is to return blank and null cells.
         *  {@link MissingCellPolicy}
         */
        public MissingCellPolicy MissingCellPolicy
        {
            get
            {
                return _missingCellPolicy;
            }
            set
            {
                _missingCellPolicy = value;
            }
        }

        /**
         * Validate sheet index
         *
         * @param index the index to validate
         * @throws ArgumentException if the index is out of range (index
         *            &lt; 0 || index &gt;= NumberOfSheets).
        */
        private void ValidateSheetIndex(int index)
        {
            int lastSheetIx = sheets.Count - 1;
            if (index < 0 || index > lastSheetIx)
            {
                String range = "(0.." + lastSheetIx + ")";
                if (lastSheetIx == -1)
                {
                    range = "(no sheets)";
                }
                throw new ArgumentException("Sheet index ("
                        + index + ") is out of range " + range);
            }
        }

        /**
         * Gets the first tab that is displayed in the list of tabs in excel.
         *
         * @return integer that Contains the index to the active sheet in this book view.
         */
        public int FirstVisibleTab
        {
            get
            {
                CT_BookViews bookViews = workbook.bookViews;
                CT_BookView bookView = bookViews.GetWorkbookViewArray(0);
                return (int)bookView.firstSheet;
            }
            set
            {
                CT_BookViews bookViews = workbook.bookViews;
                CT_BookView bookView = bookViews.GetWorkbookViewArray(0);
                bookView.firstSheet = (uint)value;
            }
        }

        /**
         * Sets the printarea for the sheet provided
         * <p>
         * i.e. Reference = $A$1:$B$2
         * @param sheetIndex Zero-based sheet index (0 Represents the first sheet to keep consistent with java)
         * @param reference Valid name Reference for the Print Area
         */
        public void SetPrintArea(int sheetIndex, String reference)
        {
            XSSFName name = GetBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
            if (name == null)
            {
                name = CreateBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
            }
            //short externSheetIndex = Workbook.CheckExternSheet(sheetIndex);
            //name.SetExternSheetNumber(externSheetIndex);
            String[] parts = COMMA_PATTERN.Split(reference);
            StringBuilder sb = new StringBuilder(32);
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0)
                {
                    sb.Append(",");
                }
                SheetNameFormatter.AppendFormat(sb, GetSheetName(sheetIndex));
                sb.Append("!");
                sb.Append(parts[i]);
            }
            name.RefersToFormula = (sb.ToString());
        }

        /**
         * For the Convenience of Java Programmers maintaining pointers.
         * @see #setPrintArea(int, String)
         * @param sheetIndex Zero-based sheet index (0 = First Sheet)
         * @param startColumn Column to begin printarea
         * @param endColumn Column to end the printarea
         * @param startRow Row to begin the printarea
         * @param endRow Row to end the printarea
         */
        public void SetPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow)
        {
            String reference = GetReferencePrintArea(GetSheetName(sheetIndex), startColumn, endColumn, startRow, endRow);
            SetPrintArea(sheetIndex, reference);
        }


        private static String GetReferenceBuiltInRecord(String sheetName, int startC, int endC, int startR, int endR)
        {
            //windows excel example for built-in title: 'second sheet'!$E:$F,'second sheet'!$2:$3
            CellReference colRef = new CellReference(sheetName, 0, startC, true, true);
            CellReference colRef2 = new CellReference(sheetName, 0, endC, true, true);

            String escapedName = SheetNameFormatter.Format(sheetName);

            String c;
            if (startC == -1 && endC == -1) c = "";
            else c = escapedName + "!$" + colRef.CellRefParts[2] + ":$" + colRef2.CellRefParts[2];

            CellReference rowRef = new CellReference(sheetName, startR, 0, true, true);
            CellReference rowRef2 = new CellReference(sheetName, endR, 0, true, true);

            String r = "";
            if (startR == -1 && endR == -1) r = "";
            else
            {
                if (!rowRef.CellRefParts[1].Equals("0") && !rowRef2.CellRefParts[1].Equals("0"))
                {
                    r = escapedName + "!$" + rowRef.CellRefParts[1] + ":$" + rowRef2.CellRefParts[1];
                }
            }

            using var rng = ZString.CreateStringBuilder();
            rng.Append(c);
            if (rng.Length > 0 && r.Length > 0) rng.Append(',');
            rng.Append(r);
            return rng.ToString();
        }

        private static String GetReferencePrintArea(String sheetName, int startC, int endC, int startR, int endR)
        {
            //windows excel example: Sheet1!$C$3:$E$4
            CellReference colRef = new CellReference(sheetName, startR, startC, true, true);
            CellReference colRef2 = new CellReference(sheetName, endR, endC, true, true);

            return "$" + colRef.CellRefParts[2] + "$" + colRef.CellRefParts[1] + ":$" + colRef2.CellRefParts[2] + "$" + colRef2.CellRefParts[1];
        }

        public XSSFName GetBuiltInName(String builtInCode, int sheetNumber)
        {
            if (!namedRangesByName.ContainsKey(builtInCode.ToLower()))
                return null;

            foreach (XSSFName name in namedRangesByName[builtInCode.ToLower()])
            {
                if (name.SheetIndex == sheetNumber)
                {
                    return name;
                }
            }
            return null;
        }

        /**
         * Generates a NameRecord to represent a built-in region
         *
         * @return a new NameRecord
         * @throws ArgumentException if sheetNumber is invalid
         * @throws POIXMLException if such a name already exists in the workbook
         */
        internal XSSFName CreateBuiltInName(String builtInName, int sheetNumber)
        {
            ValidateSheetIndex(sheetNumber);

            CT_DefinedNames names = workbook.definedNames == null ? workbook.AddNewDefinedNames() : workbook.definedNames;
            CT_DefinedName nameRecord = names.AddNewDefinedName();
            nameRecord.name = (builtInName);
            nameRecord.localSheetId = (uint)sheetNumber;
            nameRecord.localSheetIdSpecified = true;

            if (GetBuiltInName(builtInName, sheetNumber) != null)
            {
                throw new POIXMLException("Builtin (" + builtInName
                            + ") already exists for sheet (" + sheetNumber + ")");
            }

            return CreateAndStoreName(nameRecord);
        }

        /**
         * We only Set one sheet as selected for compatibility with HSSF.
         */
        public void SetSelectedTab(int index)
        {
            int idx = 0;
            foreach (XSSFSheet sh in sheets)
            {
                sh.IsSelected = idx == index;
                idx++;
            }
        }

        /**
         * Set the sheet name.
         *
         * @param sheetIndex sheet number (0 based)
         * @param sheetname  the new sheet name
         * @throws ArgumentException if the name is null or invalid
         *  or workbook already Contains a sheet with this name
         * @see {@link #CreateSheet(String)}
         * @see {@link NPOI.ss.util.WorkbookUtil#CreateSafeSheetName(String nameProposal)}
         *      for a safe way to create valid names
         */
        public void SetSheetName(int sheetIndex, String sheetname)
        {
            ValidateSheetIndex(sheetIndex);
            String oldSheetName = GetSheetName(sheetIndex);

            WorkbookUtil.ValidateSheetName(sheetname);

            // Do nothing if no change
            if (sheetname.Equals(oldSheetName)) return;

            // Check it isn't already taken
            if (ContainsSheet(sheetname, sheetIndex))
                throw new ArgumentException(string.Format("The workbook already contains a sheet named '{0}'", sheetname));

            // Update references to the name
            XSSFFormulaUtils utils = new XSSFFormulaUtils(this);
            utils.UpdateSheetName(sheetIndex, oldSheetName, sheetname);

            workbook.sheets.GetSheetArray(sheetIndex).name = (sheetname);
        }

        /**
         * Sets the order of appearance for a given sheet.
         *
         * @param sheetname the name of the sheet to reorder
         * @param pos the position that we want to insert the sheet into (0 based)
         */
        public void SetSheetOrder(String sheetname, int pos)
        {
            int idx = GetSheetIndex(sheetname);
            XSSFSheet sheet = sheets[idx];
            sheets.RemoveAt(idx);
            sheets.Insert(pos, sheet);

            // Reorder CT_Sheets
            CT_Sheets ct = workbook.sheets;
            CT_Sheet cts = ct.GetSheetArray(idx).Copy();
            workbook.sheets.RemoveSheet(idx);
            CT_Sheet newcts = ct.InsertNewSheet(pos);
            newcts.Set(cts);

            //notify sheets
            List<CT_Sheet> sheetArray = ct.sheet;
            for (int i = 0; i < sheetArray.Count; i++)
            {
                sheets[i].sheet = sheetArray[i];
            }

            UpdateNamedRangesAfterSheetReorder(idx, pos);
            UpdateActiveSheetAfterSheetReorder(idx, pos);
        }

        /**
         * update sheet-scoped named ranges in this workbook after changing the sheet order
         * of a sheet at oldIndex to newIndex.
         * Sheets between these indices will move left or right by 1.
         *
         * @param oldIndex the original index of the re-ordered sheet
         * @param newIndex the new index of the re-ordered sheet
         */
        private void UpdateNamedRangesAfterSheetReorder(int oldIndex, int newIndex)
        {
            // update sheet index of sheet-scoped named ranges
            foreach (XSSFName name in namedRanges)
            {
                int i = name.SheetIndex;
                // name has sheet-level scope
                if (i != -1)
                {
                    // name refers to this sheet
                    if (i == oldIndex)
                    {
                        name.SheetIndex = newIndex;
                    }
                    // if oldIndex > newIndex then this sheet moved left and sheets between newIndex and oldIndex moved right
                    else if (newIndex <= i && i < oldIndex)
                    {
                        name.SheetIndex = i + 1;
                    }
                    // if oldIndex < newIndex then this sheet moved right and sheets between oldIndex and newIndex moved left
                    else if (oldIndex < i && i <= newIndex)
                    {
                        name.SheetIndex = i - 1;
                    }
                }
            }
        }

        private void UpdateActiveSheetAfterSheetReorder(int oldIndex, int newIndex)
        {
            // adjust active sheet if necessary
            int active = ActiveSheetIndex;
            if (active == oldIndex)
            {
                // moved sheet was the active one
                SetActiveSheet(newIndex);
            }
            else if ((active < oldIndex && active < newIndex) ||
                     (active > oldIndex && active > newIndex))
            {
                // not affected
            }
            else if (newIndex > oldIndex)
            {
                // moved sheet was below before and is above now => active is one less
                SetActiveSheet(active - 1);
            }
            else
            {
                // remaining case: moved sheet was higher than active before and is lower now => active is one more
                SetActiveSheet(active + 1);
            }
        }

        /**
         * marshal named ranges from the {@link #namedRanges} collection to the underlying CT_Workbook bean
         */
        private void SaveNamedRanges()
        {
            // Named ranges
            if (namedRanges.Count > 0)
            {
                CT_DefinedNames names = new CT_DefinedNames();
                List<CT_DefinedName> nr = new List<CT_DefinedName>(namedRanges.Count);
                foreach (XSSFName name in namedRanges)
                {
                    nr.Add(name.GetCTName());
                }
                names.SetDefinedNameArray(nr);
                if (workbook.IsSetDefinedNames())
                {
                    workbook.unsetDefinedNames();
                }
                workbook.SetDefinedNames(names);
                // Re-process the named ranges
                ReprocessNamedRanges();
            }
            else
            {
                if (workbook.IsSetDefinedNames())
                {
                    workbook.unsetDefinedNames();
                }
            }
        }
        private void ReprocessNamedRanges()
        {
            namedRangesByName = new Dictionary<string, List<XSSFName>>();
            namedRanges = new List<XSSFName>();
            if (workbook.IsSetDefinedNames())
            {
                foreach (CT_DefinedName ctName in workbook.definedNames.definedName)
                {
                    CreateAndStoreName(ctName);
                }
            }
        }
        private void SaveCalculationChain()
        {
            if (calcChain != null)
            {
                int count = calcChain.GetCTCalcChain().SizeOfCArray();
                if (count == 0)
                {
                    RemoveRelation(calcChain);
                    calcChain = null;
                }
            }
        }


        protected internal override void Commit()
        {
            SaveNamedRanges();
            SaveCalculationChain();

            PackagePart part = GetPackagePart();
            doc.Save(part.GetOutputStream());
        }

        /// <summary>
        /// Write the document to the specified stream, and optionally leave the stream open without closing it.
        /// </summary>
        /// <param name="stream">the stream you wish to write the xlsx to</param>
        /// <param name="leaveOpen">leave stream open or not</param>
        public void Write(Stream stream, bool leaveOpen = false)
        {
            bool? originalValue = null;
            if (Package is ZipPackage package)
            {
                //By default ZipPackage closes the stream if it wasn't constructed from a stream.
                originalValue = ((ZipPackage)Package).IsExternalStream;
                ((ZipPackage)Package).IsExternalStream = leaveOpen;
            }
            base.Write(stream);
            if (originalValue.HasValue && Package is ZipPackage)
            {
                ((ZipPackage)Package).IsExternalStream = originalValue.Value;
            }
        }

        /**
         * Returns SharedStringsTable - the cache of strings for this workbook
         *
         * @return the shared string table
         */

        public SharedStringsTable GetSharedStringSource()
        {
            return this.sharedStringSource;
        }

        /**
         * Return a object representing a collection of shared objects used for styling content,
         * e.g. fonts, cell styles, colors, etc.
         */
        public StylesTable GetStylesSource()
        {
            return this.stylesSource;
        }

        /**
         * Returns the Theme of current workbook.
         */
        public ThemesTable GetTheme()
        {
            if (stylesSource == null) return null;
            return stylesSource.GetTheme();
        }

        /**
         * Returns an object that handles instantiating concrete
         *  classes of the various instances for XSSF.
         */
        public ICreationHelper GetCreationHelper()
        {
            if (_creationHelper == null) _creationHelper = new XSSFCreationHelper(this);
            return _creationHelper;
        }

        /**
         * Determines whether a workbook Contains the provided sheet name.
         * For the purpose of comparison, long names are tRuncated to 31 chars.
         *
         * @param name the name to Test (case insensitive match)
         * @param excludeSheetIdx the sheet to exclude from the check or -1 to include all sheets in the Check.
         * @return true if the sheet Contains the name, false otherwise.
         */
        //@SuppressWarnings("deprecation") //  GetXYZArray() array accessors are deprecated
        private bool ContainsSheet(String name, int excludeSheetIdx)
        {
            List<CT_Sheet> ctSheetArray = workbook.sheets.sheet;

            if (name.Length > Max_SENSITIVE_SHEET_NAME_LEN)
            {
                name = name.Substring(0, Max_SENSITIVE_SHEET_NAME_LEN);
            }

            for (int i = 0; i < ctSheetArray.Count; i++)
            {
                String ctName = ctSheetArray[i].name;
                if (ctName.Length > Max_SENSITIVE_SHEET_NAME_LEN)
                {
                    ctName = ctName.Substring(0, Max_SENSITIVE_SHEET_NAME_LEN);
                }

                if (excludeSheetIdx != i && name.Equals(ctName, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Gets a bool value that indicates whether the date systems used in the workbook starts in 1904.
        /// The default value is false, meaning that the workbook uses the 1900 date system,
        /// where 1/1/1900 is the first day in the system.
        /// </summary>
        /// <returns>True if the date systems used in the workbook starts in 1904</returns>
        public bool IsDate1904()
        {
            CT_WorkbookPr workbookPr = workbook.workbookPr;

            if (workbookPr == null)
                return false;

            return workbookPr.date1904Specified && workbookPr.date1904;
        }

        /**
         * Get the document's embedded files.
         */
        public override List<PackagePart> GetAllEmbedds()
        {
            List<PackagePart> embedds = new List<PackagePart>();

            foreach (XSSFSheet sheet in sheets)
            {
                // Get the embeddings for the workbook
                foreach (PackageRelationship rel in sheet.GetPackagePart().GetRelationshipsByType(XSSFRelation.OLEEMBEDDINGS.Relation))
                {
                    embedds.Add(sheet.GetPackagePart().GetRelatedPart(rel));
                }
                foreach (PackageRelationship rel in sheet.GetPackagePart().GetRelationshipsByType(XSSFRelation.PACKEMBEDDINGS.Relation))
                {
                    embedds.Add(sheet.GetPackagePart().GetRelatedPart(rel));
                }
            }
            return embedds;
        }

        public bool IsHidden
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        /**
         * Check whether a sheet is hidden.
         * <p>
         * Note that a sheet could instead be Set to be very hidden, which is different
         *  ({@link #isSheetVeryHidden(int)})
         * </p>
         * @param sheetIx Number
         * @return <code>true</code> if sheet is hidden
         */
        public bool IsSheetHidden(int sheetIx)
        {
            ValidateSheetIndex(sheetIx);
            CT_Sheet ctSheet = sheets[sheetIx].sheet;
            return ctSheet.state == ST_SheetState.hidden;
        }

        /**
         * Check whether a sheet is very hidden.
         * <p>
         * This is different from the normal hidden status
         *  ({@link #isSheetHidden(int)})
         * </p>
         * @param sheetIx sheet index to check
         * @return <code>true</code> if sheet is very hidden
         */
        public bool IsSheetVeryHidden(int sheetIx)
        {
            ValidateSheetIndex(sheetIx);
            CT_Sheet ctSheet = sheets[sheetIx].sheet;
            return ctSheet.state == ST_SheetState.veryHidden;
        }

        public SheetVisibility GetSheetVisibility(int sheetIx)
        {
            ValidateSheetIndex(sheetIx);
            CT_Sheet ctSheet = sheets[sheetIx].sheet;
            ST_SheetState state = ctSheet.state;
            if(state == ST_SheetState.visible)
            {
                return SheetVisibility.Visible;
            }
            if(state == ST_SheetState.hidden)
            {
                return SheetVisibility.Hidden;
            }
            if(state == ST_SheetState.veryHidden)
            {
                return SheetVisibility.VeryHidden;
            }
            throw new ArgumentException("This should never happen");
        }

        /**
         * Sets the visible state of this sheet.
         * <p>
         *   Calling <code>setSheetHidden(sheetIndex, true)</code> is equivalent to
         *   <code>setSheetHidden(sheetIndex, Workbook.SHEET_STATE_HIDDEN)</code>.
         * <br/>
         *   Calling <code>setSheetHidden(sheetIndex, false)</code> is equivalent to
         *   <code>setSheetHidden(sheetIndex, Workbook.SHEET_STATE_VISIBLE)</code>.
         * </p>
         *
         * @param sheetIx   the 0-based index of the sheet
         * @param hidden whether this sheet is hidden
         * @see #setSheetHidden(int, int)
         */
        public void SetSheetHidden(int sheetIx, bool hidden)
        {
            SetSheetHidden(sheetIx, hidden ? SheetVisibility.Hidden : SheetVisibility.Visible);
        }

        /**
         * Hide or unhide a sheet.
         *
         * <ul>
         *  <li>0 - visible. </li>
         *  <li>1 - hidden. </li>
         *  <li>2 - very hidden.</li>
         * </ul>
         * @param sheetIx the sheet index (0-based)
         * @param state one of the following <code>Workbook</code> constants:
         *        <code>Workbook.SHEET_STATE_VISIBLE</code>,
         *        <code>Workbook.SHEET_STATE_HIDDEN</code>, or
         *        <code>Workbook.SHEET_STATE_VERY_HIDDEN</code>.
         * @throws ArgumentException if the supplied sheet index or state is invalid
         */
        [Obsolete]
        public void SetSheetHidden(int sheetIx, SheetVisibility state)
        {
            SetSheetVisibility(sheetIx, state);
        }

        /// <summary>
        /// Hide or unhide a sheet.
        /// </summary>
        /// <param name="sheetIx">The sheet number</param>
        /// <param name="hidden">0 for not hidden, 1 for hidden, 2 for very hidden</param>
        [Obsolete]
        public void SetSheetHidden(int sheetIx, int state)
        {
            WorkbookUtil.ValidateSheetState((SheetVisibility)state);
            SetSheetVisibility(sheetIx, (SheetVisibility) state);
        }

        public void SetSheetVisibility(int sheetIx, SheetVisibility visibility)
        {
            ValidateSheetIndex(sheetIx);

            CT_Sheet ctSheet = sheets[sheetIx].sheet;
            switch(visibility)
            {
                case SheetVisibility.Visible:
                    ctSheet.state = (ST_SheetState.visible);
                    break;
                case SheetVisibility.Hidden:
                    ctSheet.state = (ST_SheetState.hidden);
                    break;
                case SheetVisibility.VeryHidden:
                    ctSheet.state = (ST_SheetState.veryHidden);
                    break;
                default:
                    throw new ArgumentException("This should never happen");
            }
        }

        /**
         * Fired when a formula is deleted from this workbook,
         * for example when calling cell.SetCellFormula(null)
         *
         * @see XSSFCell#setCellFormula(String)
         */
        internal void OnDeleteFormula(XSSFCell cell)
        {
            if (calcChain != null)
            {
                int sheetId = (int)((XSSFSheet)cell.Sheet).sheet.sheetId;
                calcChain.RemoveItem(sheetId, cell.GetReference());
            }
        }

        /**
         * Return the CalculationChain object for this workbook
         * <p>
         *   The calculation chain object specifies the order in which the cells in a workbook were last calculated
         * </p>
         *
         * @return the <code>CalculationChain</code> object or <code>null</code> if not defined
         */

        public CalculationChain GetCalculationChain()
        {
            return calcChain;
        }

        /**
         * Returns the list of {@link ExternalLinksTable} object for this workbook
         * 
         * <p>The external links table specifies details of named ranges etc
         *  that are referenced from other workbooks, along with the last seen
         *  values of what they point to.</p>
         *
         * <p>Note that Excel uses index 0 for the current workbook, so the first
         *  External Links in a formula would be '[1]Foo' which corresponds to
         *  entry 0 in this list.</p>

         * @return the <code>ExternalLinksTable</code> list, which may be empty
         */
        public List<ExternalLinksTable> ExternalLinksTable
        {
            get
            {
                return externalLinks;
            }
        }

        /**
         *
         * @return a collection of custom XML mappings defined in this workbook
         */
        public List<XSSFMap> GetCustomXMLMappings()
        {
            return mapInfo == null ? new List<XSSFMap>() : mapInfo.GetAllXSSFMaps();
        }

        /**
         *
         * @return the helper class used to query the custom XML mapping defined in this workbook
         */

        public MapInfo GetMapInfo()
        {
            return mapInfo;
        }

        /**
         * Adds the External Link Table part and relations required to allow formulas 
         *  referencing the specified external workbook to be added to this one. Allows
         *  formulas such as "[MyOtherWorkbook.xlsx]Sheet3!$A$5" to be added to the 
         *  file, for workbooks not already linked / referenced.
         *
         * @param name The name the workbook will be referenced as in formulas
         * @param workbook The open workbook to fetch the link required information from
         */
        public int LinkExternalWorkbook(String name, IWorkbook workbook)
        {
            throw new RuntimeException("Not Implemented - see bug #57184");
        }
        /**
         * Specifies a bool value that indicates whether structure of workbook is locked. <br/>
         * A value true indicates the structure of the workbook is locked. Worksheets in the workbook can't be Moved,
         * deleted, hidden, unhidden, or Renamed, and new worksheets can't be inserted.<br/>
         * A value of false indicates the structure of the workbook is not locked.<br/>
         * 
         * @return true if structure of workbook is locked
         */
        public bool IsStructureLocked()
        {
            return WorkbookProtectionPresent() && workbook.workbookProtection.lockStructure;
        }

        /**
         * Specifies a bool value that indicates whether the windows that comprise the workbook are locked. <br/>
         * A value of true indicates the workbook windows are locked. Windows are the same size and position each time the
         * workbook is opened.<br/>
         * A value of false indicates the workbook windows are not locked.
         * 
         * @return true if windows that comprise the workbook are locked
         */
        public bool IsWindowsLocked()
        {
            return WorkbookProtectionPresent() && workbook.workbookProtection.lockWindows;
        }

        /**
         * Specifies a bool value that indicates whether the workbook is locked for revisions.
         * 
         * @return true if the workbook is locked for revisions.
         */
        public bool IsRevisionLocked()
        {
            return WorkbookProtectionPresent() && workbook.workbookProtection.lockRevision;
        }

        /**
         * Locks the structure of workbook.
         */
        public void LockStructure()
        {
            CreateProtectionFieldIfNotPresent();
            workbook.workbookProtection.lockStructure = (true);
        }

        /**
         * Unlocks the structure of workbook.
         */
        public void UnlockStructure()
        {
            CreateProtectionFieldIfNotPresent();
            workbook.workbookProtection.lockStructure = (false);
        }

        /**
         * Locks the windows that comprise the workbook. 
         */
        public void LockWindows()
        {
            CreateProtectionFieldIfNotPresent();
            workbook.workbookProtection.lockWindows = (true);
        }

        /**
         * Unlocks the windows that comprise the workbook. 
         */
        public void UnlockWindows()
        {
            CreateProtectionFieldIfNotPresent();
            workbook.workbookProtection.lockWindows = (false);
        }

        /**
         * Locks the workbook for revisions.
         */
        public void LockRevision()
        {
            CreateProtectionFieldIfNotPresent();
            workbook.workbookProtection.lockRevision = (true);
        }

        /**
         * Unlocks the workbook for revisions.
         */
        public void UnlockRevision()
        {
            CreateProtectionFieldIfNotPresent();
            workbook.workbookProtection.lockRevision = (false);
        }

        /**
         * Remove Pivot Tables and PivotCaches from the workbooka
         */
        public void RemovePivotTables()
        {
            foreach (var xssfPivotTable in pivotTables)
            {
                var sheet = xssfPivotTable.GetParent();
                if (sheet is XSSFSheet)
                {
                    sheet.RemoveRelation(xssfPivotTable);
                }
            }

            foreach (var poixmlDocumentPart in GetRelations())
            {
                if (poixmlDocumentPart is XSSFPivotCacheDefinition pivotCacheDefinition)
                {
                    RemoveRelation(pivotCacheDefinition);
                }
            }
        }

        private bool WorkbookProtectionPresent()
        {
            return workbook.workbookProtection != null;
        }

        private void CreateProtectionFieldIfNotPresent()
        {
            if (workbook.workbookProtection == null)
            {
                workbook.workbookProtection = (new CT_WorkbookProtection());
            }
        }

        /**
         *
         * Returns the locator of user-defined functions.
         * <p>
         * The default instance : the built-in functions with the Excel Analysis Tool Pack.
         * To Set / Evaluate custom functions you need to register them as follows:
         *
         *
         *
         * </p>
         * @return wrapped instance of UDFFinder that allows seeking functions both by index and name
         */
        /*package*/
        internal UDFFinder GetUDFFinder()
        {
            return _udfFinder;
        }

        /**
         * Register a new toolpack in this workbook.
         *
         * @param toopack the toolpack to register
         */
        public void AddToolPack(UDFFinder toopack)
        {
            _udfFinder.Add(toopack);
        }

        /**
         * Whether the application shall perform a full recalculation when the workbook is opened.
         * <p>
         * Typically you want to force formula recalculation when you modify cell formulas or values
         * of a workbook previously Created by Excel. When Set to true, this flag will tell Excel
         * that it needs to recalculate all formulas in the workbook the next time the file is opened.
         * </p>
         * <p>
         * Note, that recalculation updates cached formula results and, thus, modifies the workbook.
         * Depending on the version, Excel may prompt you with "Do you want to save the Changes in <em>filename</em>?"
         * on close.
         * </p>
         *
         * @param value true if the application will perform a full recalculation of
         * workbook values when the workbook is opened
         * @since 3.8
         */
        public void SetForceFormulaRecalculation(bool value)
        {
            CT_Workbook ctWorkbook = GetCTWorkbook();

            CT_CalcPr calcPr = ctWorkbook.IsSetCalcPr() ? ctWorkbook.calcPr : ctWorkbook.AddNewCalcPr();
            // when Set to 0, will tell Excel that it needs to recalculate all formulas
            // in the workbook the next time the file is opened.
            calcPr.calcId = 0;
            if (value && calcPr.calcMode == ST_CalcMode.manual)
            {
                calcPr.calcMode = (ST_CalcMode.auto);
            }
        }

        /**
         * Whether Excel will be asked to recalculate all formulas when the  workbook is opened.
         *
         * @since 3.8
         */
        public bool GetForceFormulaRecalculation()
        {
            CT_Workbook ctWorkbook = GetCTWorkbook();
            CT_CalcPr calcPr = ctWorkbook.calcPr;
            return calcPr != null && calcPr.calcId != 0;
        }

        /// <summary>
        /// Returns the spreadsheet version (EXCLE2007) of this workbook
        /// </summary>
        public SpreadsheetVersion SpreadsheetVersion
        {
            get
            {
                return SpreadsheetVersion.EXCEL2007;
            }
        }

        /**
         * Returns the data table with the given name (case insensitive).
         * 
         * @param name the data table name (case-insensitive)
         * @return The Data table in the workbook named <tt>name</tt>, or <tt>null</tt> if no table is named <tt>name</tt>.
         * @since 3.15 beta 2
         */
        public XSSFTable GetTable(String name)
        {
            if (name != null && sheets != null)
            {
                foreach (XSSFSheet sheet in sheets)
                {
                    foreach (XSSFTable tbl in sheet.GetTables())
                    {
                        if (name.Equals(tbl.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            return tbl;
                        }
                    }
                }
            }
            return null;
        }

        public void SetActiveSheet(int sheetIndex)
        {
            ValidateSheetIndex(sheetIndex);

            foreach (CT_BookView arrayBook in workbook.bookViews.workbookView)
            {
                arrayBook.activeTab = (uint)(sheetIndex);
            }
        }
        /**
          * Add pivotCache to the workbook
          */
        protected internal CT_PivotCache AddPivotCache(String rId)
        {
            CT_Workbook ctWorkbook = GetCTWorkbook();
            CT_PivotCaches caches;
            if (ctWorkbook.IsSetPivotCaches())
            {
                caches = ctWorkbook.pivotCaches;
            }
            else
            {
                caches = ctWorkbook.AddNewPivotCaches();
            }
            CT_PivotCache cache = caches.AddNewPivotCache();

            int tableId = PivotTables.Count + 1;
            cache.cacheId = (uint)tableId;
            cache.id = (/*setter*/rId);
            if (pivotCaches == null)
            {
                pivotCaches = new List<CT_PivotCache>();
            }
            pivotCaches.Add(cache);
            return cache;
        }

        public List<XSSFPivotTable> PivotTables
        {
            get
            {
                return pivotTables;
            }
            set
            {
                this.pivotTables = value;
            }
        }
        public bool CellFormulaValidation
        {
            get { return this.cellFormulaValidation; }
            set { this.cellFormulaValidation = false; }
        }

        #region IWorkbook Members


        public int AddPicture(byte[] pictureData, PictureType format)
        {
            int imageNumber = 1;
            List<XSSFPictureData> allPics = (List<XSSFPictureData>)GetAllPictures();

            if (allPics.Any())
            {
                List<int> sortedIndexs = new List<int> { 0 };

                sortedIndexs.AddRange
                    (
                        allPics
                            .Select(pic => XSSFPictureData.RELATIONS[(int)pic.PictureType].GetFileNameIndex(pic))
                            .OrderBy(i => i)
                            .ToList()
                    );

                int previous = sortedIndexs[0];
                for (int index = 1; index < sortedIndexs.Count; index++)
                {
                    if (sortedIndexs[index] > previous + 1)
                        break;

                    previous = sortedIndexs[index];
                }

                imageNumber = previous + 1;
            }
            
            XSSFPictureData img = (XSSFPictureData)CreateRelationship(XSSFPictureData.RELATIONS[(int)format], XSSFFactory.GetInstance(), imageNumber, true).DocumentPart;
            try
            {
                Stream out1 = img.GetPackagePart().GetOutputStream();
                out1.Write(pictureData, 0, pictureData.Length);
                out1.Close();
            }
            catch (IOException e)
            {
                throw new POIXMLException(e);
            }
            pictures.Add(img);

            // returns image Index
            return allPics.Count - 1;
        }

        public XSSFWorkbookType WorkbookType
        {
            get
            {
                return IsMacroEnabled() ? XSSFWorkbookType.XLSM : XSSFWorkbookType.XLSX;
            }
            set
            {
                try
                {
                    GetPackagePart().ContentType = (value.ContentType);
                }
                catch (InvalidFormatException e)
                {
                    throw new POIXMLException(e);
                }
            }
        }


        /**
         * Adds a vbaProject.bin file to the workbook.  This will change the workbook
         * type if necessary.
         *
         * @throws IOException
         */
        public void SetVBAProject(Stream vbaProjectStream)
        {
            if (!IsMacroEnabled())
            {
                WorkbookType = (XSSFWorkbookType.XLSM);
            }

            PackagePartName ppName;
            try
            {
                ppName = PackagingUriHelper.CreatePartName(XSSFRelation.VBA_MACROS.DefaultFileName);
            }
            catch (InvalidFormatException e)
            {
                throw new POIXMLException(e);
            }
            OPCPackage opc = Package;
            Stream outputStream;
            if (!opc.ContainPart(ppName))
            {
                POIXMLDocumentPart relationship = CreateRelationship(XSSFRelation.VBA_MACROS, XSSFFactory.GetInstance());
                outputStream = relationship.GetPackagePart().GetOutputStream();
            }
            else
            {
                PackagePart part = opc.GetPart(ppName);
                outputStream = part.GetOutputStream();
            }
            try
            {
                IOUtils.Copy(vbaProjectStream, outputStream);
            }
            finally
            {
                IOUtils.CloseQuietly(outputStream);
            }
        }

        /**
         * Adds a vbaProject.bin file taken from another, given workbook to this one.
         * @throws IOException
         * @throws InvalidFormatException
         */
        public void SetVBAProject(XSSFWorkbook macroWorkbook)
        {
            if (!macroWorkbook.IsMacroEnabled())
            {
                return;
            }
            Stream vbaProjectStream = XSSFRelation.VBA_MACROS.GetContents(macroWorkbook.CorePart);
            if (vbaProjectStream != null)
            {
                SetVBAProject(vbaProjectStream);
            }
        }

        #endregion

        #region IList<XSSFSheet> Members

        public int IndexOf(ISheet item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, ISheet item)
        {
            this.sheets.Insert(index, (XSSFSheet)item);
        }

        public void RemoveAt(int index)
        {
            this.RemoveSheetAt(index);
        }

        public ISheet this[int index]
        {
            get
            {
                return this.GetSheetAt(index);
            }
            set
            {
                if (this.sheets[index] != null)
                {
                    this.sheets[index] = (XSSFSheet)value;
                }
                else
                {
                    this.sheets.Insert(index, (XSSFSheet)value);
                }
            }
        }

        #endregion

        #region ICollection<ISheet> Members

        public void Add(ISheet item)
        {
            this.sheets.Add((XSSFSheet)item);
        }

        public void Clear()
        {
            this.sheets.Clear();
        }

        public bool Contains(ISheet item)
        {
            return this.sheets.Contains(item as XSSFSheet);
        }

        public void CopyTo(ISheet[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { return this.NumberOfSheets; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public bool Remove(ISheet item)
        {
            string sheetName = item.SheetName;
            int idx = sheets.FindIndex(_ => _.SheetName.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase));

            if (idx != -1)
            {
                RemoveSheetAt(idx);
                return true;
            }

            return false;
        }

        #endregion

        public void Dispose()
        {
            this.Close();
        }
    }
}


