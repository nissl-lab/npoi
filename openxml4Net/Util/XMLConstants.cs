
namespace NPOI.Util
{
    public static class XMLConstants
    {
        /**
         * <p>Namespace URI to use to represent that there is no Namespace.</p>
         *
         * <p>Defined by the Namespace specification to be "".</p>
         *
         * @see <a href="http://www.w3.org/TR/REC-xml-names/#defaulting">
         * Namespaces in XML, 5.2 Namespace Defaulting</a>
         */
        public static string NULL_NS_URI = "";

        /**
         * <p>Prefix to use to represent the default XML Namespace.</p>
         *
         * <p>Defined by the XML specification to be "".</p>
         *
         * @see <a
         * href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames">
         * Namespaces in XML, 3. Qualified Names</a>
         */
        public static string DEFAULT_NS_PREFIX = "";

        /**
         * <p>The official XML Namespace name URI.</p>
         *
         * <p>Defined by the XML specification to be
         * "{@code http://www.w3.org/XML/1998/namespace}".</p>
         *
         * @see <a
         * href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames">
         * Namespaces in XML, 3. Qualified Names</a>
         */
        public static string XML_NS_URI = "http://www.w3.org/XML/1998/namespace";

        /**
         * <p>The official XML Namespace prefix.</p>
         *
         * <p>Defined by the XML specification to be "{@code xml}".</p>
         *
         * @see <a
         * href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames">
         * Namespaces in XML, 3. Qualified Names<</a>
         */
        public static string XML_NS_PREFIX = "xml";

        /**
         * <p>The official XML attribute used for specifying XML Namespace
         * declarations, {@link #XMLNS_ATTRIBUTE
         * XMLConstants.XMLNS_ATTRIBUTE}, Namespace name URI.</p>
         *
         * <p>Defined by the XML specification to be
         * "{@code http://www.w3.org/2000/xmlns/}".</p>
         *
         * @see <a
         * href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames">
         * Namespaces in XML, 3. Qualified Names</a>
         * @see <a
         * href="http://www.w3.org/XML/xml-names-19990114-errata">
         * Namespaces in XML Errata</a>
         */
        public static string XMLNS_ATTRIBUTE_NS_URI = "http://www.w3.org/2000/xmlns/";

        /**
         * <p>The official XML attribute used for specifying XML Namespace
         * declarations.</p>
         *
         * <p>It is <strong><em>NOT</em></strong> valid to use as a
         * prefix.  Defined by the XML specification to be
         * "{@code xmlns}".</p>
         *
         * @see <a
         * href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames">
         * Namespaces in XML, 3. Qualified Names</a>
         */
        public static string XMLNS_ATTRIBUTE = "xmlns";

        /**
         * <p>W3C XML Schema Namespace URI.</p>
         *
         * <p>Defined to be "{@code http://www.w3.org/2001/XMLSchema}".
         *
         * @see <a href=
         *  "http://www.w3.org/TR/xmlschema-1/#Instance_Document_Constructions">
         *  XML Schema Part 1:
         *  Structures, 2.6 Schema-Related Markup in Documents Being Validated</a>
         */
        public static string W3C_XML_SCHEMA_NS_URI = "http://www.w3.org/2001/XMLSchema";

        /**
         * <p>W3C XML Schema Instance Namespace URI.</p>
         *
         * <p>Defined to be "{@code http://www.w3.org/2001/XMLSchema-instance}".</p>
         *
         * @see <a href=
         *  "http://www.w3.org/TR/xmlschema-1/#Instance_Document_Constructions">
         *  XML Schema Part 1:
         *  Structures, 2.6 Schema-Related Markup in Documents Being Validated</a>
         */
        public static string W3C_XML_SCHEMA_INSTANCE_NS_URI = "http://www.w3.org/2001/XMLSchema-instance";

        /**
         * <p>W3C XPath Datatype Namespace URI.</p>
         *
         * <p>Defined to be "{@code http://www.w3.org/2003/11/xpath-datatypes}".</p>
         *
         * @see <a href="http://www.w3.org/TR/xpath-datamodel">XQuery 1.0 and XPath 2.0 Data Model</a>
         */
        public static string W3C_XPATH_DATATYPE_NS_URI = "http://www.w3.org/2003/11/xpath-datatypes";

        /**
         * <p>XML Document Type Declaration Namespace URI as an arbitrary value.</p>
         *
         * <p>Since not formally defined by any existing standard, arbitrarily define to be "{@code http://www.w3.org/TR/REC-xml}".
         */
        public static string XML_DTD_NS_URI = "http://www.w3.org/TR/REC-xml";

        /**
         * <p>RELAX NG Namespace URI.</p>
         *
         * <p>Defined to be "{@code http://relaxng.org/ns/structure/1.0}".</p>
         *
         * @see <a href="http://relaxng.org/spec-20011203.html">RELAX NG Specification</a>
         */
        public static string RELAXNG_NS_URI = "http://relaxng.org/ns/structure/1.0";

    }
}
