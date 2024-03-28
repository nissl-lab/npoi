
namespace NPOI.Util
{
    public static class XMLConstants
    {
        /// <summary>
        /// <para>
        /// Namespace URI to use to represent that there is no Namespace.
        /// </para>
        /// <para>
        /// Defined by the Namespace specification to be "".
        /// </para>
        /// <para>
        /// <see href="http://www.w3.org/TR/REC-xml-names/#defaulting"> Namespaces in XML, 5.2 Namespace Defaulting</see>
        /// </para>
        /// </summary>
        public static string NULL_NS_URI = "";

        /// <summary>
        /// <para>
        /// Prefix to use to represent the default XML Namespace.
        /// </para>
        /// <para>
        /// Defined by the XML specification to be "".
        /// </para>
        /// <para>
        /// <see href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames"> Namespaces in XML, 3. Qualified Names</see>
        /// </para>
        /// </summary>
        public static string DEFAULT_NS_PREFIX = "";

        /// <summary>
        /// <para>
        /// The official XML Namespace name URI.
        /// </para>
        /// <para>
        /// Defined by the XML specification to be
        /// "<c>http://www.w3.org/XML/1998/namespace</c>".
        /// </para>
        /// <para>
        /// <see href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames"> Namespaces in XML, 3. Qualified Names</see>
        /// </para>
        /// </summary>
        public static string XML_NS_URI = "http://www.w3.org/XML/1998/namespace";

        /// <summary>
        /// <para>
        /// The official XML Namespace prefix.
        /// </para>
        /// <para>
        /// Defined by the XML specification to be "<c>xml</c>".
        /// </para>
        /// <para>
        /// <see href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames"> Namespaces in XML, 3. Qualified Names<</see>
        /// </para>
        /// </summary>
        public static string XML_NS_PREFIX = "xml";

        /// <summary>
        /// <para>
        /// The official XML attribute used for specifying XML Namespace
        /// declarations, {@link #XMLNS_ATTRIBUTE
        /// XMLConstants.XMLNS_ATTRIBUTE}, Namespace name URI.
        /// </para>
        /// <para>
        /// Defined by the XML specification to be
        /// "<c>http://www.w3.org/2000/xmlns/</c>".
        /// </para>
        /// <para>
        /// <see href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames"> Namespaces in XML, 3. Qualified Names</see>
        /// <see href="http://www.w3.org/XML/xml-names-19990114-errata"> Namespaces in XML Errata</see>
        /// </para>
        /// </summary>
        public static string XMLNS_ATTRIBUTE_NS_URI = "http://www.w3.org/2000/xmlns/";

        /// <summary>
        /// <para>
        /// The official XML attribute used for specifying XML Namespace
        /// declarations.
        /// </para>
        /// <para>
        /// It is <strong><em>NOT</em></strong> valid to use as a
        /// prefix.  Defined by the XML specification to be
        /// "<c>xmlns</c>".
        /// </para>
        /// <para>
        /// <see href="http://www.w3.org/TR/REC-xml-names/#ns-qualnames"> Namespaces in XML, 3. Qualified Names</see>
        /// </para>
        /// </summary>
        public static string XMLNS_ATTRIBUTE = "xmlns";

        /// <summary>
        /// <para>
        /// W3C XML Schema Namespace URI.
        /// </para>
        /// <para>
        /// Defined to be "<c>http://www.w3.org/2001/XMLSchema</c>".
        /// </para>
        /// <para>
        /// <a href=  "http://www.w3.org/TR/xmlschema-1/#Instance_Document_Constructions">  XML Schema Part 1:  Structures, 2.6 Schema-Related Markup in Documents Being Validated</a>
        /// </para>
        /// </summary>
        public static string W3C_XML_SCHEMA_NS_URI = "http://www.w3.org/2001/XMLSchema";

        /// <summary>
        /// <para>
        /// W3C XML Schema Instance Namespace URI.
        /// </para>
        /// <para>
        /// Defined to be "<c>http://www.w3.org/2001/XMLSchema-instance</c>".
        /// </para>
        /// <para>
        /// <a href=  "http://www.w3.org/TR/xmlschema-1/#Instance_Document_Constructions">  XML Schema Part 1:  Structures, 2.6 Schema-Related Markup in Documents Being Validated</a>
        /// </para>
        /// </summary>
        public static string W3C_XML_SCHEMA_INSTANCE_NS_URI = "http://www.w3.org/2001/XMLSchema-instance";

        /// <summary>
        /// <para>
        /// W3C XPath Datatype Namespace URI.
        /// </para>
        /// <para>
        /// Defined to be "<c>http://www.w3.org/2003/11/xpath-datatypes</c>".
        /// </para>
        /// <para>
        /// <see href="http://www.w3.org/TR/xpath-datamodel">XQuery 1.0 and XPath 2.0 Data Model</see>
        /// </para>
        /// </summary>
        public static string W3C_XPATH_DATATYPE_NS_URI = "http://www.w3.org/2003/11/xpath-datatypes";

        /// <summary>
        /// <para>
        /// XML Document Type Declaration Namespace URI as an arbitrary value.
        /// </para>
        /// <para>
        /// Since not formally defined by any existing standard, arbitrarily define to be "<c>http://www.w3.org/TR/REC-xml</c>".
        /// </para>
        /// </summary>
        public static string XML_DTD_NS_URI = "http://www.w3.org/TR/REC-xml";

        /// <summary>
        /// <para>
        /// RELAX NG Namespace URI.
        /// </para>
        /// <para>
        /// Defined to be "<c>http://relaxng.org/ns/structure/1.0</c>".
        /// </para>
        /// <para>
        /// <see href="http://relaxng.org/spec-20011203.html">RELAX NG Specification</see>
        /// </para>
        /// </summary>
        public static string RELAXNG_NS_URI = "http://relaxng.org/ns/structure/1.0";

    }
}
