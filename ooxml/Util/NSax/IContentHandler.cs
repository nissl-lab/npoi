namespace NSAX {
  using NSAX.Helpers;

  /// <summary>
  ///   Receive notification of the logical content of a document.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This is the main interface that most SAX applications
  ///     implement: if the application needs to be informed of basic parsing
  ///     events, it implements this interface and registers an instance with
  ///     the SAX parser using the <see cref="IXmlReader.set_ContentHandler" /> /.
  ///     method. The parser uses the instance to report
  ///     basic document-related events like the start and end of elements
  ///     and character data.
  ///   </para>
  ///   <para>
  ///     The order of events in this interface is very important, and
  ///     mirrors the order of information in the document itself.  For
  ///     example, all of an element's content (character data, processing
  ///     instructions, and/or subelements) will appear, in order, between
  ///     the startElement event and the corresponding endElement event.
  ///   </para>
  /// </summary>
  /// <seealso cref="IXmlReader" />
  /// <seealso cref="IDTDHandler" />
  /// <seealso cref="IErrorHandler" />
  public interface IContentHandler {
    /// <summary>
    ///   Receive an object for locating the origin of SAX document events.
    ///   <para>
    ///     SAX parsers are strongly encouraged (though not absolutely
    ///     required) to supply a locator: if it does so, it must supply
    ///     the locator to the application by invoking this method before
    ///     invoking any of the other methods in the ContentHandler
    ///     interface.
    ///   </para>
    ///   <para>
    ///     The locator allows the application to determine the end
    ///     position of any document-related event, even if the parser is
    ///     not reporting an error.  Typically, the application will
    ///     use this information for reporting its own errors (such as
    ///     character content that does not match an application's
    ///     business rules).  The information returned by the locator
    ///     is probably not sufficient for use with a search engine.
    ///   </para>
    ///   <para>
    ///     Note that the locator will return correct information only
    ///     during the invocation SAX event callbacks after
    ///     <see cref="StartDocument"/> returns and before
    ///     <see cref="EndDocument"/> is called.  The
    ///     application should not attempt to use it at any other time.
    ///   </para>
    /// </summary>
    /// <param name="locator">
    ///   an object that can return the location of
    ///   any SAX document event
    /// </param>
    /// <seealso cref="ILocator" />
    void SetDocumentLocator(ILocator locator);

    /// <summary>
    ///   Receive notification of the beginning of a document.
    ///   <para>
    ///     The SAX parser will invoke this method only once, before any
    ///     other event callbacks (except for <see cref="SetDocumentLocator" />).
    ///   </para>
    /// </summary>
    /// <exception cref="SAXException">any SAX exception, possibly wrapping another exception</exception>
    /// <seealso cref="EndDocument" />
    void StartDocument();

    /// <summary>
    ///   Receive notification of the end of a document.
    ///   <para>
    ///     <strong>
    ///       There is an apparent contradiction between the
    ///       documentation for this method and the documentation for <see cref="IErrorHandler.FatalError" />.
    ///       Until this ambiguity is
    ///       resolved in a future major release, clients should make no
    ///       assumptions about whether <see cref="EndDocument"/> will or will not be
    ///       invoked when the parser has reported a <see cref="IErrorHandler.FatalError"/> or thrown
    ///       an exception.
    ///     </strong>
    ///   </para>
    ///   <para>
    ///     The SAX parser will invoke this method only once, and it will
    ///     be the last method invoked during the parse.  The parser shall
    ///     not invoke this method until it has either abandoned parsing
    ///     (because of an unrecoverable error) or reached the end of
    ///     input.
    ///   </para>
    /// </summary>
    /// <exception cref="SAXException">any SAX exception, possibly wrapping another exception</exception>
    /// <seealso cref="StartDocument" />
    void EndDocument();

    /// <summary>
    ///   Begin the scope of a prefix-URI Namespace mapping.
    ///   <para>
    ///     The information from this event is not necessary for
    ///     normal Namespace processing: the SAX XML reader will
    ///     automatically replace prefixes for element and attribute
    ///     names when the <c>http://xml.org/sax/features/namespaces</c>
    ///     feature is <c>true</c> (the default).
    ///   </para>
    ///   <para>
    ///     There are cases, however, when applications need to
    ///     use prefixes in character data or in attribute values,
    ///     where they cannot safely be expanded automatically; the
    ///     start/endPrefixMapping event supplies the information
    ///     to the application to expand prefixes in those contexts
    ///     itself, if necessary.
    ///   </para>
    ///   <para>
    ///     Note that start/endPrefixMapping events are not
    ///     guaranteed to be properly nested relative to each other:
    ///     all startPrefixMapping events will occur immediately before the
    ///     corresponding <see cref="StartElement" /> event,
    ///     and all <see cref="EndPrefixMapping" />
    ///     events will occur immediately after the corresponding
    ///     <see cref="EndElement" /> event,
    ///     but their order is not otherwise
    ///     guaranteed.
    ///   </para>
    ///   <para>
    ///     There should never be start/endPrefixMapping events for the
    ///     "xml" prefix, since it is predeclared and immutable.
    ///   </para>
    /// </summary>
    /// <param name="prefix">
    ///   the Namespace prefix being declared.
    ///   An empty string is used for the default element namespace,
    ///   which has no prefix.
    /// </param>
    /// <param name="uri">
    ///   the Namespace URI the prefix is mapped to
    /// </param>
    /// <exception cref="SAXException">the client may throw an exception during processing</exception>
    /// <seealso cref="EndPrefixMapping" />
    /// <seealso cref="StartElement" />
    void StartPrefixMapping(string prefix, string uri);

    /// <summary>
    ///   End the scope of a prefix-URI mapping.
    ///   <para>
    ///     See <see cref="StartPrefixMapping" /> for
    ///     details.  These events will always occur immediately after the
    ///     corresponding <see cref="EndElement" /> event, but the order of
    ///     <see cref="EndPrefixMapping" /> events is not otherwise
    ///     guaranteed.
    ///   </para>
    /// </summary>
    /// <param name="prefix">
    ///   the prefix that was being mapped.
    ///   This is the empty string when a default mapping scope ends.
    /// </param>
    /// <exception cref="SAXException">the client may throw an exception during processing</exception>
    /// <seealso cref="StartPrefixMapping" />
    /// <seealso cref="EndElement" />
    void EndPrefixMapping(string prefix);

    /// <summary>
    ///   Receive notification of the beginning of an element.
    ///   <para>
    ///     The Parser will invoke this method at the beginning of every
    ///     element in the XML document; there will be a corresponding
    ///     <see cref="EndElement" /> event for every startElement event
    ///     (even when the element is empty). All of the element's content will be
    ///     reported, in order, before the corresponding endElement
    ///     event.
    ///   </para>
    ///   <para>
    ///     This event allows up to three name components for each
    ///     element:
    ///   </para>
    ///   <ol>
    ///     <li>the Namespace URI;</li>
    ///     <li>the local name; and</li>
    ///     <li>the qualified (prefixed) name.</li>
    ///   </ol>
    ///   <para>
    ///     Any or all of these may be provided, depending on the
    ///     values of the <c>http://xml.org/sax/features/namespaces</c>
    ///     and the <c>http://xml.org/sax/features/namespace-prefixes</c>
    ///     properties:
    ///   </para>
    ///   <ul>
    ///     <li>
    ///       the Namespace URI and local name are required when
    ///       the namespaces property is <c>true</c> (the default), and are
    ///       optional when the namespaces property is <c>false</c> (if one is
    ///       specified, both must be);
    ///     </li>
    ///     <li>
    ///       the qualified name is required when the namespace-prefixes property
    ///       is <c>true</c>, and is optional when the namespace-prefixes property
    ///       is <c>false</c> (the default).
    ///     </li>
    ///   </ul>
    ///   <para>
    ///     Note that the attribute list provided will contain only
    ///     attributes with explicit values (specified or defaulted):
    ///     #IMPLIED attributes will be omitted.  The attribute list
    ///     will contain attributes used for Namespace declarations
    ///     (xmlns* attributes) only if the
    ///     <c>http://xml.org/sax/features/namespace-prefixes</c>
    ///     property is true (it is false by default, and support for a
    ///     true value is optional).
    ///   </para>
    ///   <para>
    ///     Like <see cref="Characters" />, attribute values may have
    ///     characters that need more than one <c>char</c> value.
    ///   </para>
    /// </summary>
    /// <param name="uri">
    ///   the Namespace URI, or the empty string if the
    ///   element has no Namespace URI or if Namespace
    ///   processing is not being performed
    /// </param>
    /// <param name="localName">
    ///   the local name (without prefix), or the
    ///   empty string if Namespace processing is not being
    ///   performed
    /// </param>
    /// <param name="qName">
    ///   the qualified name (with prefix), or the
    ///   empty string if qualified names are not available
    /// </param>
    /// <param name="atts">
    ///   the attributes attached to the element.  If
    ///   there are no attributes, it shall be an empty
    ///   Attributes object.  The value of this object after
    ///   startElement returns is undefined
    /// </param>
    /// <exception cref="SAXException">any SAX exception, possibly wrapping another exception</exception>
    /// <seealso cref="EndElement" />
    /// <seealso cref="IAttributes" />
    /// <seealso cref="Attributes " />
    void StartElement(string uri, string localName, string qName, IAttributes atts);

    /// <summary>
    ///   Receive notification of the end of an element.
    ///   <para>
    ///     The SAX parser will invoke this method at the end of every
    ///     element in the XML document; there will be a corresponding
    ///     <see cref="StartElement" /> event for every endElement
    ///     event (even when the element is empty).
    ///   </para>
    ///   <para>For information on the names, see startElement.</para>
    /// </summary>
    /// <param name="uri">
    ///   the Namespace URI, or the empty string if the
    ///   element has no Namespace URI or if Namespace
    ///   processing is not being performed
    /// </param>
    /// <param name="localName">
    ///   the local name (without prefix), or the
    ///   empty string if Namespace processing is not being
    ///   performed
    /// </param>
    /// <param name="qName">
    ///   the qualified XML name (with prefix), or the
    ///   empty string if qualified names are not available
    /// </param>
    /// <exception cref="SAXException">any SAX exception, possiblywrapping another exception
    /// </exception>
    void EndElement(string uri, string localName, string qName);

    /// <summary>
    ///   Receive notification of character data.
    ///   <para>
    ///     The Parser will call this method to report each chunk of
    ///     character data.  SAX parsers may return all contiguous character
    ///     data in a single chunk, or they may split it into several
    ///     chunks; however, all of the characters in any single event
    ///     must come from the same external entity so that the Locator
    ///     provides useful information.
    ///   </para>
    ///   <para>
    ///     The application must not attempt to read from the array
    ///     outside of the specified range.
    ///   </para>
    ///   <para>
    ///     Individual characters may consist of more than one Java
    ///     <c>char</c> value.  There are two important cases where this
    ///     happens, because characters can't be represented in just sixteen bits.
    ///     In one case, characters are represented in a <em>Surrogate Pair</em>,
    ///     using two special Unicode values. Such characters are in the so-called
    ///     "Astral Planes", with a code point above U+FFFF.  A second case involves
    ///     composite characters, such as a base character combining with one or
    ///     more accent characters.
    ///   </para>
    ///   <para>
    ///     Your code should not assume that algorithms using
    ///     <c>char</c>-at-a-time idioms will be working in character
    ///     units; in some cases they will split characters.  This is relevant
    ///     wherever XML permits arbitrary characters, such as attribute values,
    ///     processing instruction data, and comments as well as in data reported
    ///     from this method.  It's also generally relevant whenever Java code
    ///     manipulates internationalized text; the issue isn't unique to XML.
    ///   </para>
    ///   <para>
    ///     Note that some parsers will report whitespace in element
    ///     content using the <see cref="IgnorableWhitespace" />
    ///     method rather than this one (validating parsers <em>must</em>
    ///     do so).
    ///   </para>
    /// </summary>
    /// <param name="ch">
    ///   the characters from the XML document
    /// </param>
    /// <param name="start">
    ///   the start position in the array
    /// </param>
    /// <param name="length">
    ///   the number of characters to read from the array
    /// </param>
    /// <exception cref="SAXException">any SAX exception, possibly wrapping another exception</exception>
    /// <seealso cref="IgnorableWhitespace" />
    /// <seealso cref="ILocator" />
    void Characters(char[] ch, int start, int length);

    /// <summary>
    ///   Receive notification of ignorable whitespace in element content.
    ///   <para>
    ///     Validating Parsers must use this method to report each chunk
    ///     of whitespace in element content (see the W3C XML 1.0
    ///     recommendation, section 2.10): non-validating parsers may also
    ///     use this method if they are capable of parsing and using
    ///     content models.
    ///   </para>
    ///   <para>
    ///     SAX parsers may return all contiguous whitespace in a single
    ///     chunk, or they may split it into several chunks; however, all of
    ///     the characters in any single event must come from the same
    ///     external entity, so that the Locator provides useful
    ///     information.
    ///   </para>
    ///   <para>
    ///     The application must not attempt to read from the array
    ///     outside of the specified range.
    ///   </para>
    /// </summary>
    /// <param name="ch">
    ///   the characters from the XML document
    /// </param>
    /// <param name="start">
    ///   the start position in the array
    /// </param>
    /// <param name="length">
    ///   the number of characters to read from the array
    /// </param>
    /// <exception cref="SAXException">any SAX exception, possibly wrapping another exception</exception>
    /// <seealso cref="Characters" />
    void IgnorableWhitespace(char[] ch, int start, int length);

    /// <summary>
    ///   Receive notification of a processing instruction.
    ///   <para>
    ///     The Parser will invoke this method once for each processing
    ///     instruction found: note that processing instructions may occur
    ///     before or after the main document element.
    ///   </para>
    ///   <para>
    ///     A SAX parser must never report an XML declaration (XML 1.0,
    ///     section 2.8) or a text declaration (XML 1.0, section 4.3.1)
    ///     using this method.
    ///   </para>
    ///   <para>
    ///     Like <see cref="Characters" />, processing instruction
    ///     data may have characters that need more than one <c>char</c>
    ///     value.
    ///   </para>
    /// </summary>
    /// <param name="target">
    ///   the processing instruction target
    /// </param>
    /// <param name="data">
    ///   the processing instruction data, or null if
    ///   none was supplied.  The data does not include any
    ///   whitespace separating it from the target
    /// </param>
    /// <exception cref="SAXException">any SAX exception, possibly wrapping another exception</exception>
    void ProcessingInstruction(string target, string data);

    ////
    /// <summary>
    ///   Receive notification of a skipped entity.
    ///   This is not called for entity references within markup constructs
    ///   such as element start tags or markup declarations.  (The XML
    ///   recommendation requires reporting skipped external entities.
    ///   SAX also reports internal entity expansion/non-expansion, except
    ///   within markup constructs.)
    ///   <para>
    ///     The Parser will invoke this method each time the entity is
    ///     skipped.  Non-validating processors may skip entities if they
    ///     have not seen the declarations (because, for example, the
    ///     entity was declared in an external DTD subset).  All processors
    ///     may skip external entities, depending on the values of the
    ///     <c>http://xml.org/sax/features/external-general-entities</c>
    ///     and the
    ///     <c>http://xml.org/sax/features/external-parameter-entities</c>
    ///     properties.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   the name of the skipped entity.  If it is a
    ///   parameter entity, the name will begin with '%', and if
    ///   it is the external DTD subset, it will be the string
    ///   "[dtd]"
    /// </param>
    /// <exception cref="SAXException">any SAX exception, possibly wrapping another exception</exception>
    void SkippedEntity(string name);
  }

  // end of ContentHandler.java
}
