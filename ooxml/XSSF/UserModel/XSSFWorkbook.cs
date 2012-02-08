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

namespace NPOI.xssf.usermodel;















using javax.xml.namespace.QName;

using NPOI.POIXMLDocument;
using NPOI.POIXMLDocumentPart;
using NPOI.POIXMLException;
using NPOI.POIXMLProperties;
using NPOI.ss.formula.SheetNameFormatter;
using NPOI.Openxml4j.exceptions.OpenXML4JException;
using NPOI.Openxml4j.opc.OPCPackage;
using NPOI.Openxml4j.opc.PackagePart;
using NPOI.Openxml4j.opc.PackagePartName;
using NPOI.Openxml4j.opc.PackageRelationship;
using NPOI.Openxml4j.opc.PackageRelationshipTypes;
using NPOI.Openxml4j.opc.PackagingURIHelper;
using NPOI.Openxml4j.opc.TargetMode;
using NPOI.ss.formula.udf.UDFFinder;
using NPOI.ss.usermodel.Row;
using NPOI.ss.usermodel.Sheet;
using NPOI.ss.usermodel.Workbook;
using NPOI.ss.usermodel.Row.MissingCellPolicy;
using NPOI.ss.util.CellReference;
using NPOI.ss.util.WorkbookUtil;
using NPOI.util.*;
using NPOI.xssf.Model.*;
using NPOI.xssf.usermodel.helpers.XSSFFormulaUtils;
using org.apache.xmlbeans.XmlException;
using org.apache.xmlbeans.XmlObject;
using org.apache.xmlbeans.XmlOptions;
using org.Openxmlformats.schemas.officeDocument.x2006.relationships.STRelationshipId;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.*;

/**
 * High level representation of a SpreadsheetML workbook.  This is the first object most users
 * will construct whether they are Reading or writing a workbook.  It is also the
 * top level object for creating new sheets/etc.
 */
public class XSSFWorkbook : POIXMLDocument : Workbook, Iterable<XSSFSheet> {
    private static Pattern COMMA_PATTERN = Pattern.compile(",");

    /**
     * Width of one character of the default font in pixels. Same for Calibry and Arial.
     */
    public static float DEFAULT_CHARACTER_WIDTH = 7.0017f;

    /**
     * Excel silently tRuncates long sheet names to 31 chars.
     * This constant is used to ensure uniqueness in the first 31 chars
     */
    private static int MAX_SENSITIVE_SHEET_NAME_LEN = 31;

    /**
     * The underlying XML bean
     */
    private CTWorkbook workbook;

    /**
     * this holds the XSSFSheet objects attached to this workbook
     */
    private List<XSSFSheet> sheets;

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

    private ThemesTable theme;

    /**
     * The locator of user-defined functions.
     * By default includes functions from the Excel Analysis Toolpack
     */
    private IndexedUDFFinder _udfFinder = new IndexedUDFFinder(UDFFinder.DEFAULT);

    /**
     * TODO
     */
    private CalculationChain calcChain;

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
    private MissingCellPolicy _missingCellPolicy = Row.RETURN_NULL_AND_BLANK;

    /**
     * array of pictures for this workbook
     */
    private List<XSSFPictureData> pictures;

    private static POILogger logger = POILogFactory.GetLogger(XSSFWorkbook.class);

    /**
     * cached instance of XSSFCreationHelper for this workbook
     * @see {@link #getCreationHelper()}
     */
    private XSSFCreationHelper _creationHelper;

    /**
     * Create a new SpreadsheetML workbook.
     */
    public XSSFWorkbook() {
        base(newPackage());
        onWorkbookCreate();
    }

    /**
     * Constructs a XSSFWorkbook object given a OpenXML4J <code>Package</code> object,
     * see <a href="http://openxml4j.org/">www.Openxml4j.org</a>.
     *
     * @param pkg the OpenXML4J <code>Package</code> object.
     */
    public XSSFWorkbook(OPCPackage pkg)  {
        base(pkg);

        //build a tree of POIXMLDocumentParts, this workbook being the root
        load(XSSFFactory.GetInstance());
    }

    public XSSFWorkbook(InputStream is)  {
        base(PackageHelper.Open(is));
        
        //build a tree of POIXMLDocumentParts, this workbook being the root
        load(XSSFFactory.GetInstance());
    }

    /**
     * Constructs a XSSFWorkbook object given a file name.
     *
     * @param      path   the file name.
     */
    public XSSFWorkbook(String path)  {
        this(openPackage(path));
    }

    
    @SuppressWarnings("deprecation") //  GetXYZArray() array accessors are deprecated
    protected void onDocumentRead()  {
        try {
            WorkbookDocument doc = WorkbookDocument.Factory.Parse(getPackagePart().GetInputStream());
            this.workbook = doc.GetWorkbook();

            Dictionary<String, XSSFSheet> shIdMap = new Dictionary<String, XSSFSheet>();
            foreach(POIXMLDocumentPart p in GetRelations()){
                if(p is SharedStringsTable) sharedStringSource = (SharedStringsTable)p;
                else if(p is StylesTable) stylesSource = (StylesTable)p;
                else if(p is ThemesTable) theme = (ThemesTable)p;
                else if(p is CalculationChain) calcChain = (CalculationChain)p;
                else if(p is MapInfo) mapInfo = (MapInfo)p;
                else if (p is XSSFSheet) {
                    shIdMap.Put(p.GetPackageRelationship().GetId(), (XSSFSheet)p);
                }
            }
            stylesSource.SetTheme(theme);

            if(sharedStringSource == null) {
                //Create SST if it is missing
                sharedStringSource = (SharedStringsTable)CreateRelationship(XSSFRelation.SHARED_STRINGS, XSSFFactory.GetInstance());
            }

            // Load individual sheets. The order of sheets is defined by the order of CTSheet elements in the workbook
            sheets = new List<XSSFSheet>(shIdMap.Count);
            foreach (CTSheet ctSheet in this.workbook.GetSheets().GetSheetArray()) {
                XSSFSheet sh = shIdMap.Get(ctSheet.GetId());
                if(sh == null) {
                    logger.log(POILogger.WARN, "Sheet with name " + ctSheet.GetName() + " and r:id " + ctSheet.GetId()+ " was defined, but didn't exist in package, skipping");
                    continue;
                }
                sh.sheet = ctSheet;
                sh.onDocumentRead();
                sheets.Add(sh);
            }

            // Process the named ranges
            namedRanges = new List<XSSFName>();
            if(workbook.IsSetDefinedNames()) {
                foreach(CTDefinedName ctName in workbook.GetDefinedNames().GetDefinedNameArray()) {
                    namedRanges.Add(new XSSFName(ctName, this));
                }
            }

        } catch (XmlException e) {
            throw new POIXMLException(e);
        }
    }

    /**
     * Create a new CTWorkbook with all values Set to default
     */
    private void onWorkbookCreate() {
        workbook = CTWorkbook.Factory.newInstance();

        // don't EVER use the 1904 date system
        CTWorkbookPr workbookPr = workbook.AddNewWorkbookPr();
        workbookPr.SetDate1904(false);

        CTBookViews bvs = workbook.AddNewBookViews();
        CTBookView bv = bvs.AddNewWorkbookView();
        bv.SetActiveTab(0);
        workbook.AddNewSheets();

        POIXMLProperties.ExtendedProperties expProps = GetProperties().GetExtendedProperties();
        expProps.GetUnderlyingProperties().SetApplication(DOCUMENT_CREATOR);

        sharedStringSource = (SharedStringsTable)CreateRelationship(XSSFRelation.SHARED_STRINGS, XSSFFactory.GetInstance());
        stylesSource = (StylesTable)CreateRelationship(XSSFRelation.STYLES, XSSFFactory.GetInstance());

        namedRanges = new List<XSSFName>();
        sheets = new List<XSSFSheet>();
    }

    /**
     * Create a new SpreadsheetML namespace and Setup the default minimal content
     */
    protected static OPCPackage newPackage() {
        try {
            OPCPackage pkg = OPCPackage.Create(new MemoryStream());
            // Main part
            PackagePartName corePartName = PackagingURIHelper.CreatePartName(XSSFRelation.WORKBOOK.GetDefaultFileName());
            // Create main part relationship
            pkg.AddRelationship(corePartName, TargetMode.INTERNAL, PackageRelationshipTypes.CORE_DOCUMENT);
            // Create main document part
            pkg.CreatePart(corePartName, XSSFRelation.WORKBOOK.GetContentType());

            pkg.GetPackageProperties().SetCreatorProperty(DOCUMENT_CREATOR);

            return pkg;
        } catch (Exception e){
            throw new POIXMLException(e);
        }
    }

    /**
     * Return the underlying XML bean
     *
     * @return the underlying CTWorkbook bean
     */
    
    public CTWorkbook GetCTWorkbook() {
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
    public int AddPicture(byte[] pictureData, int format) {
        int imageNumber = GetAllPictures().Count + 1;
        XSSFPictureData img = (XSSFPictureData)CreateRelationship(XSSFPictureData.RELATIONS[format], XSSFFactory.GetInstance(), imageNumber, true);
        try {
            OutputStream out = img.GetPackagePart().GetOutputStream();
            out.Write(pictureData);
            out.Close();
        } catch (IOException e){
            throw new POIXMLException(e);
        }
        pictures.Add(img);
        return imageNumber - 1;
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
    public int AddPicture(InputStream is, int format)  {
        int imageNumber = GetAllPictures().Count + 1;
        XSSFPictureData img = (XSSFPictureData)CreateRelationship(XSSFPictureData.RELATIONS[format], XSSFFactory.GetInstance(), imageNumber, true);
        OutputStream out = img.GetPackagePart().GetOutputStream();
        IOUtils.copy(is, out);
        out.Close();
        pictures.Add(img);
        return imageNumber - 1;
    }

    /**
     * Create an XSSFSheet from an existing sheet in the XSSFWorkbook.
     *  The Cloned sheet is a deep copy of the original.
     *
     * @return XSSFSheet representing the Cloned sheet.
     * @throws ArgumentException if the sheet index in invalid
     * @throws POIXMLException if there were errors when cloning
     */
    public XSSFSheet CloneSheet(int sheetNum) {
        validateSheetIndex(sheetNum);

        XSSFSheet srcSheet = sheets.Get(sheetNum);
        String srcName = srcSheet.GetSheetName();
        String ClonedName = GetUniqueSheetName(srcName);

        XSSFSheet ClonedSheet = CreateSheet(ClonedName);
        try {
            MemoryStream out = new MemoryStream();
            srcSheet.Write(out);
            ClonedSheet.Read(new MemoryStream(out.ToArray()));
        } catch (IOException e){
            throw new POIXMLException("Failed to clone sheet", e);
        }
        CTWorksheet ct = ClonedSheet.GetCTWorksheet();
        if(ct.IsSetDrawing()) {
            logger.log(POILogger.WARN, "Cloning sheets with Drawings is not yet supported.");
            ct.unsetDrawing();
        }
        if(ct.IsSetLegacyDrawing()) {
            logger.log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
            ct.unsetLegacyDrawing();
        }

        ClonedSheet.SetSelected(false);

        // copy sheet's relations
        List<POIXMLDocumentPart> rels = srcSheet.GetRelations();
        foreach(POIXMLDocumentPart r in rels) {
            PackageRelationship rel = r.GetPackageRelationship();
            ClonedSheet.GetPackagePart().AddRelationship(rel.GetTargetURI(), rel.GetTargetMode(),rel.GetRelationshipType());
            ClonedSheet.AddRelation(rel.GetId(), r);
        }

        return ClonedSheet;
    }

    /**
     * Generate a valid sheet name based on the existing one. Used when cloning sheets.
     *
     * @param srcName the original sheet name to
     * @return clone sheet name
     */
    private String GetUniqueSheetName(String srcName) {
        int uniqueIndex = 2;
        String baseName = srcName;
        int bracketPos = srcName.LastIndexOf('(');
        if (bracketPos > 0 && srcName\.EndsWith(")")) {
            String suffix = srcName.Substring(bracketPos + 1, srcName.Length - ")".Length);
            try {
                uniqueIndex = Int32.ParseInt(suffix.Trim());
                uniqueIndex++;
                baseName = srcName.Substring(0, bracketPos).Trim();
            } catch (NumberFormatException e) {
                // contents of brackets not numeric
            }
        }
        while (true) {
            // Try and find the next sheet name that is unique
            String index = Int32.ToString(uniqueIndex++);
            String name;
            if (baseName.Length + index.Length + 2 < 31) {
                name = baseName + " (" + index + ")";
            } else {
                name = baseName.Substring(0, 31 - index.Length - 2) + "(" + index + ")";
            }

            //If the sheet name is unique, then Set it otherwise Move on to the next number.
            if (getSheetIndex(name) == -1) {
                return name;
            }
        }
    }

    /**
     * Create a new XSSFCellStyle and add it to the workbook's style table
     *
     * @return the new XSSFCellStyle object
     */
    public XSSFCellStyle CreateCellStyle() {
        return stylesSource.CreateCellStyle();
    }

    /**
     * Returns the instance of XSSFDataFormat for this workbook.
     *
     * @return the XSSFDataFormat object
     * @see NPOI.ss.usermodel.DataFormat
     */
    public XSSFDataFormat CreateDataFormat() {
        if (formatter == null)
            formatter = new XSSFDataFormat(stylesSource);
        return formatter;
    }

    /**
     * Create a new Font and add it to the workbook's font table
     *
     * @return new font object
     */
    public XSSFFont CreateFont() {
        XSSFFont font = new XSSFFont();
        font.registerTo(stylesSource);
        return font;
    }

    public XSSFName CreateName() {
        CTDefinedName ctName = CTDefinedName.Factory.newInstance();
        ctName.SetName("");
        XSSFName name = new XSSFName(ctName, this);
        namedRanges.Add(name);
        return name;
    }

    /**
     * Create an XSSFSheet for this workbook, Adds it to the sheets and returns
     * the high level representation.  Use this to create new sheets.
     *
     * @return XSSFSheet representing the new sheet.
     */
    public XSSFSheet CreateSheet() {
        String sheetname = "Sheet" + (sheets.Count);
        int idx = 0;
        while(getSheet(sheetname) != null) {
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
     *     assert 31 == sheet.GetSheetName().Length;
     *     assert "My very long sheet name which i" == sheet.GetSheetName();
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
     * @param sheetname  sheetname to Set for the sheet.
     * @return Sheet representing the new sheet.
     * @throws ArgumentException if the name is null or invalid
     *  or workbook already Contains a sheet with this name
     * @see {@link NPOI.ss.util.WorkbookUtil#CreateSafeSheetName(String nameProposal)}
     *      for a safe way to create valid names
     */
    public XSSFSheet CreateSheet(String sheetname) {
        if (sheetname == null) {
            throw new ArgumentException("sheetName must not be null");
        }

        if (ContainsSheet( sheetname, sheets.Count ))
               throw new ArgumentException( "The workbook already Contains a sheet of this name");

        // YK: Mimic Excel and silently tRuncate sheet names longer than 31 characters
        if(sheetname.Length > 31) sheetname = sheetname.Substring(0, 31);
        WorkbookUtil.validateSheetName(sheetname);

        CTSheet sheet = AddSheet(sheetname);

        int sheetNumber = 1;
        foreach(XSSFSheet sh in sheets) sheetNumber = (int)Math.max(sh.sheet.GetSheetId() + 1, sheetNumber);

        XSSFSheet wrapper = (XSSFSheet)CreateRelationship(XSSFRelation.WORKSHEET, XSSFFactory.GetInstance(), sheetNumber);
        wrapper.sheet = sheet;
        sheet.SetId(wrapper.GetPackageRelationship().GetId());
        sheet.SetSheetId(sheetNumber);
        if(sheets.Count == 0) wrapper.SetSelected(true);
        sheets.Add(wrapper);
        return wrapper;
    }

    protected XSSFDialogsheet CreateDialogsheet(String sheetname, CTDialogsheet dialogsheet) {
        XSSFSheet sheet = CreateSheet(sheetname);
        return new XSSFDialogsheet(sheet);
    }

    private CTSheet AddSheet(String sheetname) {
        CTSheet sheet = workbook.GetSheets().AddNewSheet();
        sheet.SetName(sheetname);
        return sheet;
    }

    /**
     * Finds a font that matches the one with the supplied attributes
     */
    public XSSFFont FindFont(short boldWeight, short color, short fontHeight, String name, bool italic, bool strikeout, short typeOffset, byte underline) {
        return stylesSource.FindFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
    }

    /**
     * Convenience method to Get the active sheet.  The active sheet is is the sheet
     * which is currently displayed when the workbook is viewed in Excel.
     * 'Selected' sheet(s) is a distinct concept.
     */
    public int GetActiveSheetIndex() {
        //activeTab (Active Sheet Index) Specifies an unsignedInt
        //that Contains the index to the active sheet in this book view.
        return (int)workbook.GetBookViews().GetWorkbookViewArray(0).GetActiveTab();
    }

    /**
     * Gets all pictures from the Workbook.
     *
     * @return the list of pictures (a list of {@link XSSFPictureData} objects.)
     * @see #AddPicture(byte[], int)
     */
    public List<XSSFPictureData> GetAllPictures() {
        if(pictures == null) {
            //In OOXML pictures are referred to in sheets,
            //dive into sheet's relations, select Drawings and their images
            pictures = new List<XSSFPictureData>();
            foreach(XSSFSheet sh in sheets){
                foreach(POIXMLDocumentPart dr in sh.GetRelations()){
                    if(dr is XSSFDrawing){
                        foreach(POIXMLDocumentPart img in dr.GetRelations()){
                            if(img is XSSFPictureData){
                                pictures.Add((XSSFPictureData)img);
                            }
                        }
                    }
                }
            }
        }
        return pictures;
    }

    /**
     * gGet the cell style object at the given index
     *
     * @param idx  index within the Set of styles
     * @return XSSFCellStyle object at the index
     */
    public XSSFCellStyle GetCellStyleAt(short idx) {
        return stylesSource.GetStyleAt(idx);
    }

    /**
     * Get the font at the given index number
     *
     * @param idx  index number
     * @return XSSFFont at the index
     */
    public XSSFFont GetFontAt(short idx) {
        return stylesSource.GetFontAt(idx);
    }

    public XSSFName GetName(String name) {
        int nameIndex = GetNameIndex(name);
        if (nameIndex < 0) {
            return null;
        }
        return namedRanges.Get(nameIndex);
    }

    public XSSFName GetNameAt(int nameIndex) {
        int nNames = namedRanges.Count;
        if (nNames < 1) {
            throw new InvalidOperationException("There are no defined names in this workbook");
        }
        if (nameIndex < 0 || nameIndex > nNames) {
            throw new ArgumentException("Specified name index " + nameIndex
                    + " is outside the allowable range (0.." + (nNames-1) + ").");
        }
        return namedRanges.Get(nameIndex);
    }

    /**
     * Gets the named range index by his name
     * <i>Note:</i>Excel named ranges are case-insensitive and
     * this method performs a case-insensitive search.
     *
     * @param name named range name
     * @return named range index
     */
    public int GetNameIndex(String name) {
        int i = 0;
        foreach(XSSFName nr in namedRanges) {
            if(nr.GetNameName().Equals(name)) {
                return i;
            }
            i++;
        }
        return -1;
    }

    /**
     * Get the number of styles the workbook Contains
     *
     * @return count of cell styles
     */
    public short GetNumCellStyles() {
        return (short) (stylesSource).GetNumCellStyles();
    }

    /**
     * Get the number of fonts in the this workbook
     *
     * @return number of fonts
     */
    public short GetNumberOfFonts() {
        return (short)stylesSource.GetFonts().Count;
    }

    /**
     * Get the number of named ranges in the this workbook
     *
     * @return number of named ranges
     */
    public int GetNumberOfNames() {
        return namedRanges.Count;
    }

    /**
     * Get the number of worksheets in the this workbook
     *
     * @return number of worksheets
     */
    public int GetNumberOfSheets() {
        return sheets.Count;
    }

    /**
     * Retrieves the reference for the printarea of the specified sheet, the sheet name is Appended to the reference even if it was not specified.
     * @param sheetIndex Zero-based sheet index (0 Represents the first sheet to keep consistent with java)
     * @return String Null if no print area has been defined
     */
    public String GetPrintArea(int sheetIndex) {
        XSSFName name = GetBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
        if (name == null) return null;
        //Adding one here because 0 indicates a global named region; doesnt make sense for print areas
        return name.GetRefersToFormula();

    }

    /**
     * Get sheet with the given name (case insensitive match)
     *
     * @param name of the sheet
     * @return XSSFSheet with the name provided or <code>null</code> if it does not exist
     */
    public XSSFSheet GetSheet(String name) {
        foreach (XSSFSheet sheet in sheets) {
            if (name.EqualsIgnoreCase(sheet.GetSheetName())) {
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
     *            &lt; 0 || index &gt;= GetNumberOfSheets()).
     */
    public XSSFSheet GetSheetAt(int index) {
        validateSheetIndex(index);
        return sheets.Get(index);
    }

    /**
     * Returns the index of the sheet by his name (case insensitive match)
     *
     * @param name the sheet name
     * @return index of the sheet (0 based) or <tt>-1</tt if not found
     */
    public int GetSheetIndex(String name) {
        for (int i = 0 ; i < sheets.Count ; ++i) {
            XSSFSheet sheet = sheets.Get(i);
            if (name.EqualsIgnoreCase(sheet.GetSheetName())) {
                return i;
            }
        }
        return -1;
    }

    /**
     * Returns the index of the given sheet
     *
     * @param sheet the sheet to look up
     * @return index of the sheet (0 based). <tt>-1</tt> if not found
     */
    public int GetSheetIndex(Sheet sheet) {
        int idx = 0;
        foreach(XSSFSheet sh in sheets){
            if(sh == sheet) return idx;
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
    public String GetSheetName(int sheetIx) {
        validateSheetIndex(sheetIx);
        return sheets.Get(sheetIx).GetSheetName();
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
    public Iterator<XSSFSheet> iterator() {
        return sheets.iterator();
    }
    /**
     * Are we a normal workbook (.xlsx), or a
     *  macro enabled workbook (.xlsm)?
     */
    public bool IsMacroEnabled() {
        return GetPackagePart().GetContentType().Equals(XSSFRelation.MACROS_WORKBOOK.GetContentType());
    }

    public void RemoveName(int nameIndex) {
        namedRanges.Remove(nameIndex);
    }

    public void RemoveName(String name) {
        for (int i = 0; i < namedRanges.Count; i++) {
            XSSFName nm = namedRanges.Get(i);
            if(nm.GetNameName().EqualsIgnoreCase(name)) {
                RemoveName(i);
                return;
            }
        }
        throw new ArgumentException("Named range was not found: " + name);
    }

    /**
     * Delete the printarea for the sheet specified
     *
     * @param sheetIndex 0-based sheet index (0 = First Sheet)
     */
    public void RemovePrintArea(int sheetIndex) {
        int cont = 0;
        foreach (XSSFName name in namedRanges) {
            if (name.GetNameName().Equals(XSSFName.BUILTIN_PRINT_AREA) && name.GetSheetIndex() == sheetIndex) {
                namedRanges.Remove(cont);
                break;
            }
            cont++;
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
    public void RemoveSheetAt(int index) {
        validateSheetIndex(index);

        onSheetDelete(index);

        XSSFSheet sheet = GetSheetAt(index);
        RemoveRelation(sheet);
        sheets.Remove(index);
    }

    /**
     * Gracefully remove references to the sheet being deleted
     *
     * @param index the 0-based index of the sheet to delete
     */
    private void onSheetDelete(int index) {
        //delete the CTSheet reference from workbook.xml
        workbook.GetSheets().RemoveSheet(index);

        //calculation chain is auxiliary, remove it as it may contain orphan references to deleted cells
        if(calcChain != null) {
            RemoveRelation(calcChain);
            calcChain = null;
        }

        //adjust indices of names ranges
        for (Iterator<XSSFName> it = namedRanges.iterator(); it.HasNext();) {
            XSSFName nm = it.next();
            CTDefinedName ct = nm.GetCTName();
            if(!ct.IsSetLocalSheetId()) continue;
            if (ct.GetLocalSheetId() == index) {
                it.Remove();
            } else if (ct.GetLocalSheetId() > index){
                // Bump down by one, so still points at the same sheet
                ct.SetLocalSheetId(ct.GetLocalSheetId()-1);
            }
        }
    }

    /**
     * Retrieves the current policy on what to do when
     *  Getting missing or blank cells from a row.
     * The default is to return blank and null cells.
     *  {@link MissingCellPolicy}
     */
    public MissingCellPolicy GetMissingCellPolicy() {
        return _missingCellPolicy;
    }
    /**
     * Sets the policy on what to do when
     *  Getting missing or blank cells from a row.
     * This will then apply to all calls to
     *  {@link Row#getCell(int)}}. See
     *  {@link MissingCellPolicy}
     */
    public void SetMissingCellPolicy(MissingCellPolicy missingCellPolicy) {
        _missingCellPolicy = missingCellPolicy;
    }

    /**
     * Convenience method to Set the active sheet.  The active sheet is is the sheet
     * which is currently displayed when the workbook is viewed in Excel.
     * 'Selected' sheet(s) is a distinct concept.
     */
    @SuppressWarnings("deprecation") //YK: GetXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public void SetActiveSheet(int index) {

        validateSheetIndex(index);

        foreach (CTBookView arrayBook in workbook.GetBookViews().GetWorkbookViewArray()) {
            arrayBook.SetActiveTab(index);
        }
    }

    /**
     * Validate sheet index
     *
     * @param index the index to validate
     * @throws ArgumentException if the index is out of range (index
     *            &lt; 0 || index &gt;= GetNumberOfSheets()).
    */
    private void validateSheetIndex(int index) {
        int lastSheetIx = sheets.Count - 1;
        if (index < 0 || index > lastSheetIx) {
            throw new ArgumentException("Sheet index ("
                    + index +") is out of range (0.." +    lastSheetIx + ")");
        }
    }

    /**
     * Gets the first tab that is displayed in the list of tabs in excel.
     *
     * @return integer that Contains the index to the active sheet in this book view.
     */
    public int GetFirstVisibleTab() {
        CTBookViews bookViews = workbook.GetBookViews();
        CTBookView bookView = bookViews.GetWorkbookViewArray(0);
        return (short) bookView.GetActiveTab();
    }

    /**
     * Sets the first tab that is displayed in the list of tabs in excel.
     *
     * @param index integer that Contains the index to the active sheet in this book view.
     */
    public void SetFirstVisibleTab(int index) {
        CTBookViews bookViews = workbook.GetBookViews();
        CTBookView bookView= bookViews.GetWorkbookViewArray(0);
        bookView.SetActiveTab(index);
    }

    /**
     * Sets the printarea for the sheet provided
     * <p>
     * i.e. Reference = $A$1:$B$2
     * @param sheetIndex Zero-based sheet index (0 Represents the first sheet to keep consistent with java)
     * @param reference Valid name Reference for the Print Area
     */
    public void SetPrintArea(int sheetIndex, String reference) {
        XSSFName name = GetBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
        if (name == null) {
            name = CreateBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
        }
        //short externSheetIndex = GetWorkbook().CheckExternSheet(sheetIndex);
        //name.SetExternSheetNumber(externSheetIndex);
        String[] parts = COMMA_PATTERN.split(reference);
        StringBuilder sb = new StringBuilder(32);
        for (int i = 0; i < parts.Length; i++) {
            if(i>0) {
                sb.Append(",");
            }
            SheetNameFormatter.AppendFormat(sb, GetSheetName(sheetIndex));
            sb.Append("!");
            sb.Append(parts[i]);
        }
        name.SetRefersToFormula(sb.ToString());
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
    public void SetPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) {
        String reference=getReferencePrintArea(getSheetName(sheetIndex), startColumn, endColumn, startRow, endRow);
        SetPrintArea(sheetIndex, reference);
    }


    /**
     * Sets the repeating rows and columns for a sheet.
     * <p/>
     * To Set just repeating columns:
     * <pre>
     *  workbook.SetRepeatingRowsAndColumns(0,0,1,-1,-1);
     * </pre>
     * To Set just repeating rows:
     * <pre>
     *  workbook.SetRepeatingRowsAndColumns(0,-1,-1,0,4);
     * </pre>
     * To remove all repeating rows and columns for a sheet.
     * <pre>
     *  workbook.SetRepeatingRowsAndColumns(0,-1,-1,-1,-1);
     * </pre>
     *
     * @param sheetIndex  0 based index to sheet.
     * @param startColumn 0 based start of repeating columns.
     * @param endColumn   0 based end of repeating columns.
     * @param startRow    0 based start of repeating rows.
     * @param endRow      0 based end of repeating rows.
     */
    public void SetRepeatingRowsAndColumns(int sheetIndex,
                                           int startColumn, int endColumn,
                                           int startRow, int endRow) {
        //    Check arguments
        if ((startColumn == -1 && endColumn != -1) || startColumn < -1 || endColumn < -1 || startColumn > endColumn)
            throw new ArgumentException("Invalid column range specification");
        if ((startRow == -1 && endRow != -1) || startRow < -1 || endRow < -1 || startRow > endRow)
            throw new ArgumentException("Invalid row range specification");

        XSSFSheet sheet = GetSheetAt(sheetIndex);
        bool removingRange = startColumn == -1 && endColumn == -1 && startRow == -1 && endRow == -1;

        XSSFName name = GetBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
        if (removingRange) {
            if(name != null)namedRanges.Remove(name);
            return;
        }
        if (name == null) {
            name = CreateBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
        }

        String reference = GetReferenceBuiltInRecord(name.GetSheetName(), startColumn, endColumn, startRow, endRow);
        name.SetRefersToFormula(reference);

        // If the print Setup isn't currently defined, then add it
        //  in but without printer defaults
        // If it's already there, leave it as-is!
        CTWorksheet ctSheet = sheet.GetCTWorksheet();
        if(ctSheet.IsSetPageSetup() && ctSheet.IsSetPageMargins()) {
           // Everything we need is already there
        } else {
           // Have Initial ones Put in place
           XSSFPrintSetup printSetup = sheet.GetPrintSetup();
           printSetup.SetValidSettings(false);
        }
    }

    private static String GetReferenceBuiltInRecord(String sheetName, int startC, int endC, int startR, int endR) {
        //windows excel example for built-in title: 'second sheet'!$E:$F,'second sheet'!$2:$3
        CellReference colRef = new CellReference(sheetName, 0, startC, true, true);
        CellReference colRef2 = new CellReference(sheetName, 0, endC, true, true);

        String escapedName = SheetNameFormatter.format(sheetName);

        String c;
        if(startC == -1 && endC == -1) c= "";
        else c = escapedName + "!$" + colRef.GetCellRefParts()[2] + ":$" + colRef2.GetCellRefParts()[2];

        CellReference rowRef = new CellReference(sheetName, startR, 0, true, true);
        CellReference rowRef2 = new CellReference(sheetName, endR, 0, true, true);

        String r = "";
        if(startR == -1 && endR == -1) r = "";
        else {
            if (!rowRef.GetCellRefParts()[1].Equals("0") && !rowRef2.GetCellRefParts()[1].Equals("0")) {
                r = escapedName + "!$" + rowRef.GetCellRefParts()[1] + ":$" + rowRef2.GetCellRefParts()[1];
            }
        }

        StringBuilder rng = new StringBuilder();
        rng.Append(c);
        if(rng.Length > 0 && r.Length > 0) rng.Append(',');
        rng.Append(r);
        return rng.ToString();
    }

    private static String GetReferencePrintArea(String sheetName, int startC, int endC, int startR, int endR) {
        //windows excel example: Sheet1!$C$3:$E$4
        CellReference colRef = new CellReference(sheetName, startR, startC, true, true);
        CellReference colRef2 = new CellReference(sheetName, endR, endC, true, true);

        return "$" + colRef.GetCellRefParts()[2] + "$" + colRef.GetCellRefParts()[1] + ":$" + colRef2.GetCellRefParts()[2] + "$" + colRef2.GetCellRefParts()[1];
    }

    XSSFName GetBuiltInName(String builtInCode, int sheetNumber) {
        foreach (XSSFName name in namedRanges) {
            if (name.GetNameName().EqualsIgnoreCase(builtInCode) && name.GetSheetIndex() == sheetNumber) {
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
    XSSFName CreateBuiltInName(String builtInName, int sheetNumber) {
        validateSheetIndex(sheetNumber);

        CTDefinedNames names = workbook.GetDefinedNames() == null ? workbook.AddNewDefinedNames() : workbook.GetDefinedNames();
        CTDefinedName nameRecord = names.AddNewDefinedName();
        nameRecord.SetName(builtInName);
        nameRecord.SetLocalSheetId(sheetNumber);

        XSSFName name = new XSSFName(nameRecord, this);
        foreach (XSSFName nr in namedRanges) {
            if (nr.Equals(name))
                throw new POIXMLException("Builtin (" + builtInName
                        + ") already exists for sheet (" + sheetNumber + ")");
        }

        namedRanges.Add(name);
        return name;
    }

    /**
     * We only Set one sheet as selected for compatibility with HSSF.
     */
    public void SetSelectedTab(int index) {
        for (int i = 0 ; i < sheets.Count ; ++i) {
            XSSFSheet sheet = sheets.Get(i);
            sheet.SetSelected(i == index);
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
    public void SetSheetName(int sheetIndex, String sheetname) {
        validateSheetIndex(sheetIndex);

        // YK: Mimic Excel and silently tRuncate sheet names longer than 31 characters
        if(sheetname != null && sheetname.Length > 31) sheetname = sheetname.Substring(0, 31);
        WorkbookUtil.validateSheetName(sheetname);

        if (ContainsSheet(sheetname, sheetIndex ))
            throw new ArgumentException( "The workbook already Contains a sheet of this name" );

        XSSFFormulaUtils utils = new XSSFFormulaUtils(this);
        utils.updateSheetName(sheetIndex, sheetname);

        workbook.GetSheets().GetSheetArray(sheetIndex).SetName(sheetname);
    }

    /**
     * Sets the order of appearance for a given sheet.
     *
     * @param sheetname the name of the sheet to reorder
     * @param pos the position that we want to insert the sheet into (0 based)
     */
    public void SetSheetOrder(String sheetname, int pos) {
        int idx = GetSheetIndex(sheetname);
        sheets.Add(pos, sheets.Remove(idx));
        // Reorder CTSheets
        CTSheets ct = workbook.GetSheets();
        XmlObject cts = ct.GetSheetArray(idx).copy();
        workbook.GetSheets().RemoveSheet(idx);
        CTSheet newcts = ct.insertNewSheet(pos);
        newcts.Set(cts);

        //notify sheets
        for(int i=0; i < sheets.Count; i++) {
            sheets.Get(i).sheet = ct.GetSheetArray(i);
        }
    }

    /**
     * marshal named ranges from the {@link #namedRanges} collection to the underlying CTWorkbook bean
     */
    private void saveNamedRanges(){
        // Named ranges
        if(namedRanges.Count > 0) {
            CTDefinedNames names = CTDefinedNames.Factory.newInstance();
            CTDefinedName[] nr = new CTDefinedName[namedRanges.Count];
            int i = 0;
            foreach(XSSFName name in namedRanges) {
                nr[i] = name.GetCTName();
                i++;
            }
            names.SetDefinedNameArray(nr);
            workbook.SetDefinedNames(names);
        } else {
            if(workbook.IsSetDefinedNames()) {
                workbook.unsetDefinedNames();
            }
        }
    }

    private void saveCalculationChain(){
        if(calcChain != null){
            int count = calcChain.GetCTCalcChain().sizeOfCArray();
            if(count == 0){
                RemoveRelation(calcChain);
                calcChain = null;
            }
        }
    }

    
    protected void Commit()  {
        saveNamedRanges();
        saveCalculationChain();

        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        xmlOptions.SetSaveSyntheticDocumentElement(new QName(CTWorkbook.type.GetName().GetNamespaceURI(), "workbook"));
        Dictionary<String, String> map = new Dictionary<String, String>();
        map.Put(STRelationshipId.type.GetName().GetNamespaceURI(), "r");
        xmlOptions.SetSaveSuggestedPrefixes(map);

        PackagePart part = GetPackagePart();
        OutputStream out = part.GetOutputStream();
        workbook.save(out, xmlOptions);
        out.Close();
    }

    /**
     * Returns SharedStringsTable - tha cache of string for this workbook
     *
     * @return the shared string table
     */
    
    public SharedStringsTable GetSharedStringSource() {
        return this.sharedStringSource;
    }

    /**
     * Return a object representing a collection of shared objects used for styling content,
     * e.g. fonts, cell styles, colors, etc.
     */
    public StylesTable GetStylesSource() {
        return this.stylesSource;
    }

    /**
     * Returns the Theme of current workbook.
     */
    public ThemesTable GetTheme() {
        return theme;
    }

    /**
     * Returns an object that handles instantiating concrete
     *  classes of the various instances for XSSF.
     */
    public XSSFCreationHelper GetCreationHelper() {
        if(_creationHelper == null) _creationHelper = new XSSFCreationHelper(this);
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
    @SuppressWarnings("deprecation") //  GetXYZArray() array accessors are deprecated
    private bool ContainsSheet(String name, int excludeSheetIdx) {
        CTSheet[] ctSheetArray = workbook.GetSheets().GetSheetArray();

        if (name.Length > MAX_SENSITIVE_SHEET_NAME_LEN) {
            name = name.Substring(0, MAX_SENSITIVE_SHEET_NAME_LEN);
        }

        for (int i = 0; i < ctSheetArray.Length; i++) {
            String ctName = ctSheetArray[i].GetName();
            if (ctName.Length > MAX_SENSITIVE_SHEET_NAME_LEN) {
                ctName = ctName.Substring(0, MAX_SENSITIVE_SHEET_NAME_LEN);
            }

            if (excludeSheetIdx != i && name.EqualsIgnoreCase(ctName))
                return true;
        }
        return false;
    }

    /**
     * Gets a bool value that indicates whether the date systems used in the workbook starts in 1904.
     * <p>
     * The default value is false, meaning that the workbook uses the 1900 date system,
     * where 1/1/1900 is the first day in the system..
     * </p>
     * @return true if the date systems used in the workbook starts in 1904
     */
    protected bool IsDate1904(){
        CTWorkbookPr workbookPr = workbook.GetWorkbookPr();
        return workbookPr != null && workbookPr.GetDate1904();
    }

    /**
     * Get the document's embedded files.
     */
    public List<PackagePart> GetAllEmbedds() throws OpenXML4JException {
        List<PackagePart> embedds = new LinkedList<PackagePart>();

        foreach(XSSFSheet sheet in sheets){
            // Get the embeddings for the workbook
            foreach(PackageRelationship rel in sheet.GetPackagePart().GetRelationshipsByType(XSSFRelation.OLEEMBEDDINGS.GetRelation()))
                embedds.Add(getTargetPart(rel));

            foreach(PackageRelationship rel in sheet.GetPackagePart().GetRelationshipsByType(XSSFRelation.PACKEMBEDDINGS.GetRelation()))
                embedds.Add(getTargetPart(rel));

        }
        return embedds;
    }

    public bool IsHidden() {
        throw new RuntimeException("Not implemented yet");
    }

    public void SetHidden(bool hiddenFlag) {
        throw new RuntimeException("Not implemented yet");
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
    public bool IsSheetHidden(int sheetIx) {
        validateSheetIndex(sheetIx);
        CTSheet ctSheet = sheets.Get(sheetIx).sheet;
        return ctSheet.GetState() == STSheetState.HIDDEN;
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
    public bool IsSheetVeryHidden(int sheetIx) {
        validateSheetIndex(sheetIx);
        CTSheet ctSheet = sheets.Get(sheetIx).sheet;
        return ctSheet.GetState() == STSheetState.VERY_HIDDEN;
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
    public void SetSheetHidden(int sheetIx, bool hidden) {
        SetSheetHidden(sheetIx, hidden ? SHEET_STATE_HIDDEN : SHEET_STATE_VISIBLE);
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
    public void SetSheetHidden(int sheetIx, int state) {
        validateSheetIndex(sheetIx);
        WorkbookUtil.validateSheetState(state);
        CTSheet ctSheet = sheets.Get(sheetIx).sheet;
        ctSheet.SetState(STSheetState.Enum.forInt(state + 1));
    }

    /**
     * Fired when a formula is deleted from this workbook,
     * for example when calling cell.SetCellFormula(null)
     *
     * @see XSSFCell#setCellFormula(String)
     */
    protected void onDeleteFormula(XSSFCell cell){
        if(calcChain != null) {
            int sheetId = (int)cell.Sheet.sheet.GetSheetId();
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
    
    public CalculationChain GetCalculationChain(){
        return calcChain;
    }

    /**
     *
     * @return a collection of custom XML mappings defined in this workbook
     */
    public Collection<XSSFMap> GetCustomXMLMappings(){
        return mapInfo == null ? new List<XSSFMap>() : mapInfo.GetAllXSSFMaps();
    }

    /**
     *
     * @return the helper class used to query the custom XML mapping defined in this workbook
     */
    
    public MapInfo GetMapInfo(){
    	return mapInfo;
    }


	/**
	 * Specifies a bool value that indicates whether structure of workbook is locked. <br/>
	 * A value true indicates the structure of the workbook is locked. Worksheets in the workbook can't be Moved,
	 * deleted, hidden, unhidden, or Renamed, and new worksheets can't be inserted.<br/>
	 * A value of false indicates the structure of the workbook is not locked.<br/>
	 * 
	 * @return true if structure of workbook is locked
	 */
	public bool IsStructureLocked() {
		return workbookProtectionPresent() && workbook.GetWorkbookProtection().GetLockStructure();
	}

	/**
	 * Specifies a bool value that indicates whether the windows that comprise the workbook are locked. <br/>
	 * A value of true indicates the workbook windows are locked. Windows are the same size and position each time the
	 * workbook is opened.<br/>
	 * A value of false indicates the workbook windows are not locked.
	 * 
	 * @return true if windows that comprise the workbook are locked
	 */
	public bool IsWindowsLocked() {
		return workbookProtectionPresent() && workbook.GetWorkbookProtection().GetLockWindows();
	}

	/**
	 * Specifies a bool value that indicates whether the workbook is locked for revisions.
	 * 
	 * @return true if the workbook is locked for revisions.
	 */
	public bool IsRevisionLocked() {
		return workbookProtectionPresent() && workbook.GetWorkbookProtection().GetLockRevision();
	}
	
	/**
	 * Locks the structure of workbook.
	 */
	public void lockStructure() {
		CreateProtectionFieldIfNotPresent();
		workbook.GetWorkbookProtection().SetLockStructure(true);
	}
	
	/**
	 * Unlocks the structure of workbook.
	 */
	public void unLockStructure() {
		CreateProtectionFieldIfNotPresent();
		workbook.GetWorkbookProtection().SetLockStructure(false);
	}

	/**
	 * Locks the windows that comprise the workbook. 
	 */
	public void lockWindows() {
		CreateProtectionFieldIfNotPresent();
		workbook.GetWorkbookProtection().SetLockWindows(true);
	}
	
	/**
	 * Unlocks the windows that comprise the workbook. 
	 */
	public void unLockWindows() {
		CreateProtectionFieldIfNotPresent();
		workbook.GetWorkbookProtection().SetLockWindows(false);
	}
	
	/**
	 * Locks the workbook for revisions.
	 */
	public void lockRevision() {
		CreateProtectionFieldIfNotPresent();
		workbook.GetWorkbookProtection().SetLockRevision(true);
	}

	/**
	 * Unlocks the workbook for revisions.
	 */
	public void unLockRevision() {
		CreateProtectionFieldIfNotPresent();
		workbook.GetWorkbookProtection().SetLockRevision(false);
	}
	
	private bool workbookProtectionPresent() {
		return workbook.GetWorkbookProtection() != null;
	}

	private void CreateProtectionFieldIfNotPresent() {
		if (workbook.GetWorkbookProtection() == null){
			workbook.SetWorkbookProtection(CTWorkbookProtection.Factory.newInstance());
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
    /*package*/ UDFFinder GetUDFFinder() {
        return _udfFinder;
    }

    /**
     * Register a new toolpack in this workbook.
     *
     * @param toopack the toolpack to register
     */
    public void AddToolPack(UDFFinder toopack){
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
   public void SetForceFormulaRecalculation(bool value){
        CTWorkbook ctWorkbook = GetCTWorkbook();
        CTCalcPr calcPr = ctWorkbook.IsSetCalcPr() ? ctWorkbook.GetCalcPr() : ctWorkbook.AddNewCalcPr();
        // when Set to 0, will tell Excel that it needs to recalculate all formulas
        // in the workbook the next time the file is opened.
        calcPr.SetCalcId(0);
    }

    /**
     * Whether Excel will be asked to recalculate all formulas when the  workbook is opened.
     *
     * @since 3.8
     */
    public bool GetForceFormulaRecalculation(){
        CTWorkbook ctWorkbook = GetCTWorkbook();
        CTCalcPr calcPr = ctWorkbook.GetCalcPr();
        return calcPr != null && calcPr.GetCalcId() != 0;
    }

}


