/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XSSF.EventUserModel
{

    using NPOI;
    using NPOI.OpenXml4Net.Exceptions;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.UserModel;
    using NSAX;
    using NSAX.Helpers;
    using System.Xml;

    /// <summary>
    /// This class makes it easy to Get at individual parts
    ///  of an OOXML .xlsx file, suitable for low memory sax
    ///  parsing or similar.
    /// It makes up the core part of the EventUserModel support
    ///  for XSSF.
    /// </summary>
    public class XSSFReader
    {

        private static ISet<String> WORKSHEET_RELS =
                new HashSet<String>(
                        Arrays.AsList(new String[]{
                            XSSFRelation.WORKSHEET.Relation,
                            XSSFRelation.CHARTSHEET.Relation,
                        })
                );
        //private static  POILogger LOGGER = POILogFactory.GetLogger(XSSFReader.class);

        protected OPCPackage pkg;
        protected PackagePart workbookPart;

        /// <summary>
        /// Creates a new XSSFReader, for the given package
        /// </summary>
        public XSSFReader(OPCPackage pkg)
        {

            this.pkg = pkg;

            PackageRelationship coreDocRelationship = this.pkg.GetRelationshipsByType(
                    PackageRelationshipTypes.CORE_DOCUMENT).GetRelationship(0);

            // strict OOXML likely not fully supported, see #57699
            // this code is similar to POIXMLDocumentPart.PartFromOPCPackage, but I could not combine it
            // easily due to different return values
            if (coreDocRelationship == null)
            {
                if (this.pkg.GetRelationshipsByType(
                        PackageRelationshipTypes.STRICT_CORE_DOCUMENT).GetRelationship(0) != null)
                {
                    throw new POIXMLException("Strict OOXML isn't currently supported, please see bug #57699");
                }

                throw new POIXMLException("OOXML file structure broken/invalid - no core document found!");
            }

            // Get the part that holds the workbook
            workbookPart = this.pkg.GetPart(coreDocRelationship);
        }


        /// <summary>
        /// Opens up the Shared Strings Table, parses it, and
        ///  returns a handy object for working with
        ///  shared strings.
        /// </summary>
        public SharedStringsTable SharedStringsTable
        {
            get
            {
                List<PackagePart> parts = pkg.GetPartsByContentType(XSSFRelation.SHARED_STRINGS.ContentType);
                return parts.Count == 0 ? null : new SharedStringsTable(parts[0]);
            }
        }

        /// <summary>
        /// Opens up the Styles Table, parses it, and
        ///  returns a handy object for working with cell styles
        /// </summary>
        public StylesTable StylesTable
        {
            get
            {
                List<PackagePart> parts = pkg.GetPartsByContentType(XSSFRelation.STYLES.ContentType);
                if(parts.Count == 0)
                    return null;

                // Create the Styles Table, and associate the Themes if present
                StylesTable styles = new StylesTable(parts[0]);
                parts = pkg.GetPartsByContentType(XSSFRelation.THEME.ContentType);
                if(parts.Count != 0)
                {
                    styles.Theme = (new ThemesTable(parts[0]));
                }
                return styles;
            }
            
        }


        /// <summary>
        /// Returns an InputStream to read the contents of the
        ///  shared strings table.
        /// </summary>
        public Stream SharedStringsData => XSSFRelation.SHARED_STRINGS.GetContents(workbookPart);

        /// <summary>
        /// Returns an InputStream to read the contents of the
        ///  styles table.
        /// </summary>
        public Stream StylesData => XSSFRelation.STYLES.GetContents(workbookPart);

        /// <summary>
        /// Returns an InputStream to read the contents of the
        ///  themes table.
        /// </summary>
        public Stream ThemesData => XSSFRelation.THEME.GetContents(workbookPart);

        /// <summary>
        /// Returns an InputStream to read the contents of the
        ///  main Workbook, which contains key overall data for
        ///  the file, including sheet definitions.
        /// </summary>
        public Stream WorkbookData => workbookPart.GetInputStream();

        /// <summary>
        /// Returns an InputStream to read the contents of the
        ///  specified Sheet.
        /// </summary>
        /// <param name="relId">The relationId of the sheet, from a r:id on the workbook</param>
        public Stream GetSheet(String relId)
        {

            PackageRelationship rel = workbookPart.GetRelationship(relId);
            if (rel == null)
            {
                throw new ArgumentException("No Sheet found with r:id " + relId);
            }

            PackagePartName relName = PackagingUriHelper.CreatePartName(rel.TargetUri);
            PackagePart sheet = pkg.GetPart(relName);
            if (sheet == null)
            {
                throw new ArgumentException("No data found for Sheet with r:id " + relId);
            }
            return sheet.GetInputStream();
        }

        /// <summary>
        /// Returns an Iterator which will let you Get at all the
        ///  different Sheets in turn.
        /// Each sheet's InputStream is only opened when fetched
        ///  from the Iterator. It's up to you to close the
        ///  InputStreams when done with each one.
        /// </summary>
        public IEnumerator<Stream> GetSheetsData()
        {

            return new SheetIterator(workbookPart);
        }

        /// <summary>
        /// Iterator over sheet data.
        /// </summary>
        public class SheetIterator : IEnumerator<Stream>
        {

            /// <summary>
            ///  Maps relId and the corresponding PackagePart
            /// </summary>
            private Dictionary<String, PackagePart> sheetMap;

            /// <summary>
            /// Current sheet reference
            /// </summary>
            XSSFSheetRef xssfSheetRef;

            /// <summary>
            /// Iterator over CTSheet objects, returns sheets in <tt>logical</tt> order.
            /// We can't rely on the Ooxml4J's relationship iterator because it returns objects in physical order,
            /// i.e. as they are stored in the underlying package
            /// </summary>
            IEnumerator<XSSFSheetRef> sheetIterator;

            
            /// <summary>
            /// Construct a new SheetIterator
            /// </summary>
            /// <param name="wb">package part holding workbook.xml</param>
            internal SheetIterator(PackagePart wb)
            {
                /*
                 * The order of sheets is defined by the order of CTSheet elements in workbook.xml
                 */
                try
                {
                    //step 1. Map sheet's relationship Id and the corresponding PackagePart
                    sheetMap = new Dictionary<String, PackagePart>();
                    OPCPackage pkg = wb.Package;
                    ISet<String> worksheetRels = SheetRelationships;
                    foreach (PackageRelationship rel in wb.Relationships)
                    {
                        String relType = rel.RelationshipType;
                        if (worksheetRels.Contains(relType))
                        {
                            PackagePartName relName = PackagingUriHelper.CreatePartName(rel.TargetUri);
                            sheetMap.Add(rel.Id, pkg.GetPart(relName));
                        }
                    }
                    //step 2. Read array of CTSheet elements, wrap it in a LinkedList
                    //and construct an iterator
                    sheetIterator = CreateSheetIteratorFromWB(wb);
                }
                catch (InvalidFormatException e)
                {
                    throw new POIXMLException(e);
                }
            }

            IEnumerator<XSSFSheetRef> CreateSheetIteratorFromWB(PackagePart wb)
            {
                XMLSheetRefReader xmlSheetRefReader = new XMLSheetRefReader();
                NSAX.AElfred.SAXDriver xmlReader;
                try
                {
                    xmlReader = new NSAX.AElfred.SAXDriver();// SAXHelper.newXMLReader();
                }
                //catch (ParserConfigurationException e)
                //{
                //    throw new POIXMLException(e);
                //}
                catch (SAXException e)
                {
                    throw new POIXMLException(e);
                }
                xmlReader.ContentHandler = (xmlSheetRefReader);
                try
                {
                    xmlReader.Parse(new InputSource(wb.GetInputStream()));
                }
                catch (SAXException e)
                {
                    throw new POIXMLException(e);
                }

                List<XSSFSheetRef> validSheets = new List<XSSFSheetRef>();
                foreach (XSSFSheetRef xssfSheetRef in xmlSheetRefReader.GetSheetRefs())
                {
                    //if there's no relationship id, silently skip the sheet
                    String sheetId = xssfSheetRef.Id;
                    if (sheetId != null && sheetId.Length > 0)
                    {
                        validSheets.Add(xssfSheetRef);
                    }
                }
                return validSheets.GetEnumerator();
            }

            /// <summary>
            /// Gets string representations of relationships
            /// that are sheet-like.  Added to allow subclassing
            /// by XSSFBReader.  This is used to decide what
            /// relationships to load into the sheetRefs
            /// </summary>
            /// <return>all relationships that are sheet-like</return>
            ISet<String> SheetRelationships => WORKSHEET_RELS;

            /// <summary>
            /// Returns <tt>true</tt> if the iteration has more elements.
            /// </summary>
            /// <return><tt>true</tt> if the iterator has more elements.</return>
            //public bool HasNext()
            //{
            //    return sheetIterator.HasNext();
            //}

            /// <summary>
            /// Returns input stream of the next sheet in the iteration
            /// </summary>
            /// <return>input stream of the next sheet in the iteration</return>
            private Stream Next()
            {
                xssfSheetRef = sheetIterator.Current;

                String sheetId = xssfSheetRef.Id;
                try
                {
                    PackagePart sheetPkg = sheetMap[sheetId];
                    return sheetPkg.GetInputStream();
                }
                catch (IOException e)
                {
                    throw new POIXMLException(e);
                }
            }

            public Stream Current => Next();

            object IEnumerator.Current => Next();
            public bool MoveNext()
            {
                return sheetIterator.MoveNext();
            }

            public void Reset()
            {
                sheetIterator.Reset();
            }

            public void Dispose()
            {
                sheetIterator.Dispose();
            }

            /// <summary>
            /// Returns name of the current sheet
            /// </summary>
            /// <return>name of the current sheet</return>
            public String SheetName => xssfSheetRef.Name;

            /// <summary>
            /// Returns the comments associated with this sheet,
            ///  or null if there aren't any
            /// </summary>
            public CommentsTable SheetComments
            {
                get
                {
                    PackagePart sheetPkg = SheetPart;

                    // Do we have a comments relationship? (Only ever one if so)
                    try
                    {
                        PackageRelationshipCollection commentsList =
                         sheetPkg.GetRelationshipsByType(XSSFRelation.SHEET_COMMENTS.Relation);
                        if(commentsList.Size > 0)
                        {
                            PackageRelationship comments = commentsList.GetRelationship(0);
                            PackagePartName commentsName = PackagingUriHelper.CreatePartName(comments.TargetUri);
                            PackagePart commentsPart = sheetPkg.Package.GetPart(commentsName);
                            return new CommentsTable(commentsPart);
                        }
                    }
                    catch(InvalidFormatException)
                    {
                        return null;
                    }
                    catch(IOException)
                    {
                        return null;
                    }
                    return null;
                }
            }

            /// <summary>
            /// Returns the shapes associated with this sheet,
            /// an empty list or null if there is an exception
            /// </summary>
            public List<XSSFShape> Shapes
            {
                get
                {
                    PackagePart sheetPkg = SheetPart;
                    List<XSSFShape> shapes = new List<XSSFShape>();
                    // Do we have a comments relationship? (Only ever one if so)
                    try
                    {
                        PackageRelationshipCollection drawingsList = sheetPkg.GetRelationshipsByType(XSSFRelation.DRAWINGS.Relation);
                        for(int i = 0; i < drawingsList.Size; i++)
                        {
                            PackageRelationship drawings = drawingsList.GetRelationship(i);
                            PackagePartName drawingsName = PackagingUriHelper.CreatePartName(drawings.TargetUri);
                            PackagePart drawingsPart = sheetPkg.Package.GetPart(drawingsName);
                            if(drawingsPart == null)
                            {
                                //parts can go missing; Excel ignores them silently -- TIKA-2134
                                //LOGGER.log(POILogger.WARN, "Missing Drawing: " + drawingsName + ". Skipping it.");
                                continue;
                            }
                            XSSFDrawing drawing = new XSSFDrawing(drawingsPart);
                            shapes.AddRange(drawing.GetShapes());
                        }
                    }
                    catch(XmlException)
                    {
                        return null;
                    }
                    catch(InvalidFormatException)
                    {
                        return null;
                    }
                    catch(IOException)
                    {
                        return null;
                    }
                    return shapes;
                }
            }

            public PackagePart SheetPart => sheetMap[xssfSheetRef.Id];

            /// <summary>
            /// We're read only, so remove isn't supported
            /// </summary>
            public void Remove()
            {
                throw new InvalidOperationException("Not supported");
            }

           
        }

        public sealed class XSSFSheetRef
        {
            //do we need to store sheetId, too?
            private String id;
            private String name;

            public XSSFSheetRef(String id, String name)
            {
                this.id = id;
                this.name = name;
            }

            public String Id => id;

            public String Name => name;
        }

        //scrapes sheet reference info and order from workbook.xml
        private class XMLSheetRefReader : DefaultHandler
        {
            private static String SHEET = "sheet";
            private static String ID = "id";
            private static String NAME = "name";

            private List<XSSFSheetRef> sheetRefs = new List<XSSFSheetRef>();

            // read <sheet name="Sheet6" sheetId="4" r:id="rId6"/>
            // and add XSSFSheetRef(id="rId6", name="Sheet6") to sheetRefs
            public override void StartElement(String uri, String localName, String qName, IAttributes attrs)
            {

                if (localName.Equals(SHEET, StringComparison.OrdinalIgnoreCase))
                {
                    String name = null;
                    String id = null;
                    for (int i = 0; i < attrs.Length; i++)
                    {
                        String attrName = attrs.GetLocalName(i);
                        if (attrName.Equals(NAME, StringComparison.OrdinalIgnoreCase))
                        {
                            name = attrs.GetValue(i);
                        }
                        else if (attrName.Equals(ID, StringComparison.OrdinalIgnoreCase))
                        {
                            id = attrs.GetValue(i);
                        }
                        if (name != null && id != null)
                        {
                            sheetRefs.Add(new XSSFSheetRef(id, name));
                            break;
                        }
                    }
                }
            }

            public List<XSSFSheetRef> GetSheetRefs()
            {
                return sheetRefs;
            }
        }
    }
}

