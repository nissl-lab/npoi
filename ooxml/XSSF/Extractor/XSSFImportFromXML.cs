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

namespace NPOI.xssf.extractor;






using javax.xml.namespace.NamespaceContext;
using javax.xml.Parsers.DocumentBuilder;
using javax.xml.Parsers.DocumentBuilderFactory;
using javax.xml.Parsers.ParserConfigurationException;
using javax.xml.xpath.XPath;
using javax.xml.xpath.XPathConstants;
using javax.xml.xpath.XPathExpressionException;
using javax.xml.xpath.XPathFactory;

using NPOI.util.POILogFactory;
using NPOI.util.POILogger;
using NPOI.xssf.usermodel.XSSFTable;
using NPOI.xssf.usermodel.XSSFCell;
using NPOI.xssf.usermodel.XSSFMap;
using NPOI.xssf.usermodel.XSSFRow;
using NPOI.xssf.usermodel.helpers.XSSFSingleXmlCell;
using NPOI.xssf.usermodel.helpers.XSSFXmlColumnPr;
using org.w3c.dom.Document;
using org.w3c.dom.Element;
using org.w3c.dom.NamedNodeMap;
using org.w3c.dom.Node;
using org.w3c.dom.NodeList;
using org.xml.sax.InputSource;
using org.xml.sax.SAXException;

/**
 * Imports data from an external XML to an XLSX according to one of the mappings
 * defined.The output XML Schema must respect this limitations:
 * <ul>
 * <li>the input XML must be valid according to the XML Schema used in the mapping</li>
 * <li>denormalized table mapping is not supported (see OpenOffice part 4: chapter 3.5.1.7)</li>
 * <li>all the namespaces used in the document must be declared in the root node</li>
 * </ul>
 */
public class XSSFImportFromXML {

    private XSSFMap _map;

    private static POILogger logger = POILogFactory.GetLogger(XSSFImportFromXML.class);

    public XSSFImportFromXML(XSSFMap map) {
        _map = map;
    }

    /**
     * Imports an XML into the XLSX using the Custom XML mapping defined
     *
     * @param xmlInputString the XML to import
     * @throws SAXException if error occurs during XML parsing
     * @throws XPathExpressionException if error occurs during XML navigation
     * @throws ParserConfigurationException if there are problems with XML Parser configuration
     * @throws IOException  if there are problems Reading the input string
     */
    public void importFromXML(String xmlInputString) , XPathExpressionException, ParserConfigurationException, IOException {

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.SetNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();

        Document doc = builder.Parse(new InputSource(new StringReader(xmlInputString.Trim())));

        List<XSSFSingleXmlCell> SingleXmlCells = _map.GetRelatedSingleXMLCell();

        List<XSSFTable> tables = _map.GetRelatedTables();

        XPathFactory xpathFactory = XPathFactory.newInstance();
        XPath xpath = xpathFactory.newXPath();

        // Setting namespace context to XPath
        // Assuming that the namespace prefix in the mapping xpath is the
        // same as the one used in the document
        xpath.SetNamespaceContext(new DefaultNamespaceContext(doc));

        foreach (XSSFSingleXmlCell SingleXmlCell in SingleXmlCells) {

            String xpathString = SingleXmlCell.GetXpath();
            Node result = (Node) xpath.Evaluate(xpathString, doc, XPathConstants.NODE);
            String textContent = result.GetTextContent();
            logger.log(POILogger.DEBUG, "Extracting with xpath " + xpathString + " : value is '" + textContent + "'");
            XSSFCell cell = SingleXmlCell.GetReferencedCell();
            logger.log(POILogger.DEBUG, "Setting '" + textContent + "' to cell " + cell.ColumnIndex + "-" + cell.RowIndex + " in sheet "
                                            + cell.Sheet.GetSheetName());
            cell.SetCellValue(textContent);
        }

        foreach (XSSFTable table in tables) {

            String commonXPath = table.GetCommonXpath();
            NodeList result = (NodeList) xpath.Evaluate(commonXPath, doc, XPathConstants.NODESET);
            int rowOffset = table.GetStartCellReference().Row + 1;// the first row Contains the table header
            int columnOffset = table.GetStartCellReference().Col - 1;

            for (int i = 0; i < result.GetLength(); i++) {

                // TODO: implement support for denormalized XMLs (see
                // OpenOffice part 4: chapter 3.5.1.7)

                foreach (XSSFXmlColumnPr xmlColumnPr in table.GetXmlColumnPrs()) {

                    int localColumnId = (int) xmlColumnPr.GetId();
                    int rowId = rowOffset + i;
                    int columnId = columnOffset + localColumnId;
                    String localXPath = xmlColumnPr.GetLocalXPath();
                    localXPath = localXPath.Substring(localXPath.substring(1).indexOf('/') + 1);

                    // Build an XPath to select the right node (assuming
                    // that the commonXPath != "/")
                    String nodeXPath = commonXPath + "[" + (i + 1) + "]" + localXPath;

                    // TODO: convert the data to the cell format
                    String value = (String) xpath.Evaluate(nodeXPath, result.item(i), XPathConstants.STRING);
                    logger.log(POILogger.DEBUG, "Extracting with xpath " + nodeXPath + " : value is '" + value + "'");
                    XSSFRow row = table.GetXSSFSheet().GetRow(rowId);
                    if (row == null) {
                        row = table.GetXSSFSheet().CreateRow(rowId);
                    }

                    XSSFCell cell = row.GetCell(columnId);
                    if (cell == null) {
                        cell = row.CreateCell(columnId);
                    }
                    logger.log(POILogger.DEBUG, "Setting '" + value + "' to cell " + cell.ColumnIndex + "-" + cell.RowIndex + " in sheet "
                                                    + table.GetXSSFSheet().GetSheetName());
                    cell.SetCellValue(value.Trim());
                }
            }
        }
    }

    private static class DefaultNamespaceContext : NamespaceContext {
        /**
         * Node from which to start searching for a xmlns attribute that binds a
         * prefix to a namespace.
         */
        private Element _docElem;

        public DefaultNamespaceContext(Document doc) {
            _docElem = doc.GetDocumentElement();
        }

        public String GetNamespaceURI(String prefix) {
            return GetNamespaceForPrefix(prefix);
        }

        /**
         * @param prefix Prefix to Resolve.
         * @return uri of Namespace that prefix Resolves to, or
         *         <code>null</code> if specified prefix is not bound.
         */
        private String GetNamespaceForPrefix(String prefix) {

            // Code adapted from Xalan's org.apache.xml.utils.PrefixResolverDefault.GetNamespaceForPrefix()

            if (prefix.Equals("xml")) {
                return "http://www.w3.org/XML/1998/namespace";
            }

            Node parent = _docElem;

            while (parent != null) {

                int type = parent.GetNodeType();
                if (type == Node.ELEMENT_NODE) {
                    if (parent.GetNodeName().startsWith(prefix + ":")) {
                        return parent.GetNamespaceURI();
                    }
                    NamedNodeMap nnm = parent.GetAttributes();

                    for (int i = 0; i < nnm.GetLength(); i++) {
                        Node attr = nnm.item(i);
                        String aname = attr.GetNodeName();
                        bool IsPrefix = aname.startsWith("xmlns:");

                        if (isPrefix || aname.Equals("xmlns")) {
                            int index = aname.IndexOf(':');
                            String p = isPrefix ? aname.Substring(index + 1) : "";

                            if (p.Equals(prefix)) {
                                return attr.GetNodeValue();
                            }
                        }
                    }
                } else if (type == Node.ENTITY_REFERENCE_NODE) {
                    continue;
                } else {
                    break;
                }
                parent = parent.GetParentNode();
            }

            return null;
        }

        // Dummy implementation - not used!
        public Iterator GetPrefixes(String val) {
            return null;
        }

        // Dummy implementation - not used!
        public String GetPrefix(String uri) {
            return null;
        }
    }
}

