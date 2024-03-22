namespace NSAX {
  using System.IO;

  /// <summary>
  ///   Interface for reading an XML document using callbacks.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     <strong>Note:</strong> despite its name, this interface does
  ///     <em>not</em> extend the standard .net <see cref="TextReader" /> interface,
  ///     because reading XML is a fundamentally different activity
  ///     than reading character data.
  ///   </para>
  ///   <para>
  ///     XMLReader is the interface that an XML parser's SAX2 driver must
  ///     implement.  This interface allows an application to set and
  ///     query features and properties in the parser, to register
  ///     event handlers for document processing, and to initiate
  ///     a document parse.
  ///   </para>
  ///   <para>
  ///     All SAX interfaces are assumed to be synchronous: the
  ///     <see cref="Parse(string)" /> methods must not return until parsing
  ///     is complete, and readers must wait for an event-handler callback
  ///     to return before reporting the next event.
  ///   </para>
  ///   <ol>
  ///     <li>
  ///       it adds a standard way to query and set features and
  ///       properties; and
  ///     </li>
  ///     <li>
  ///       it adds Namespace support, which is required for many
  ///       higher-level XML standards.
  ///     </li>
  ///   </ol>
  /// </summary>
  /// <seealso cref="IXmlFilter" />
  public interface IXmlReader {
    /// <summary>
    ///   Gets or sets the current entity resolver.
    ///   <para>
    ///     If the application does not register an entity resolver,
    ///     the XMLReader will perform its own default resolution.
    ///   </para>
    ///   <para>
    ///     Applications may register a new or different resolver in the
    ///     middle of a parse, and the SAX parser must begin using the new
    ///     resolver immediately.
    ///   </para>
    /// </summary>
    /// <returns>The current entity resolver, or null if none has been registered.</returns>
    IEntityResolver EntityResolver { get; set; }

    /// <summary>
    ///   Gets or sets the current DTD handler.
    ///   <para>
    ///     If the application does not register a DTD handler, all DTD
    ///     events reported by the SAX parser will be silently ignored.
    ///   </para>
    ///   <para>
    ///     Applications may register a new or different handler in the
    ///     middle of a parse, and the SAX parser must begin using the new
    ///     handler immediately.
    ///   </para>
    /// </summary>
    /// <returns> The current DTD handler, or null if none has been registered.</returns>
    IDTDHandler DTDHandler { get; set; }

    /// <summary>
    ///   Gets or sets the current content handler.
    ///   <para>
    ///     If the application does not register a content handler, all
    ///     content events reported by the SAX parser will be silently
    ///     ignored.
    ///   </para>
    ///   <para>
    ///     Applications may register a new or different handler in the
    ///     middle of a parse, and the SAX parser must begin using the new
    ///     handler immediately.
    ///   </para>
    /// </summary>
    /// <returns>The current content handler, or null if none has been registered.</returns>
    IContentHandler ContentHandler { get; set; }

    /// <summary>
    ///   Gets or sets the current error handler.
    ///   <para>
    ///     If the application does not register an error handler, all
    ///     error events reported by the SAX parser will be silently
    ///     ignored; however, normal processing may not continue.  It is
    ///     highly recommended that all SAX applications implement an
    ///     error handler to avoid unexpected bugs.
    ///   </para>
    ///   <para>
    ///     Applications may register a new or different handler in the
    ///     middle of a parse, and the SAX parser must begin using the new
    ///     handler immediately.
    ///   </para>
    /// </summary>
    /// <returns>The current error handler, or null if none  has been registered.</returns>
    IErrorHandler ErrorHandler { get; set; }

    /// <summary>
    ///   Look up the value of a feature flag.
    ///   <para>
    ///     The feature name is any fully-qualified URI.  It is
    ///     possible for an XMLReader to recognize a feature name but
    ///     temporarily be unable to return its value.
    ///     Some feature values may be available only in specific
    ///     contexts, such as before, during, or after a parse.
    ///     Also, some feature values may not be programmatically accessible.
    ///     There is no implementation-independent way to expose whether the underlying
    ///     parser is performing validation, expanding external entities,
    ///     and so forth.)
    ///   </para>
    ///   <para>
    ///     All XMLReaders are required to recognize the
    ///     http://xml.org/sax/features/namespaces and the
    ///     http://xml.org/sax/features/namespace-prefixes feature names.
    ///   </para>
    ///   <para>
    ///     Implementors are free (and encouraged) to invent their own features,
    ///     using names built on their own URIs.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   The feature name, which is a fully-qualified URI.
    /// </param>
    /// <returns>The current value of the feature (true or false).</returns>
    /// <exception cref="SAXNotRecognizedException">If the feature value can't be assigned or retrieved.</exception>
    /// <exception cref="SAXNotSupportedException">
    ///   When the XMLReader recognizes the feature name but cannot
    ///   determine its value at this time.
    /// </exception>
    /// <seealso cref="SetFeature" />
    bool GetFeature(string name);

    /// <summary>
    ///   Set the value of a feature flag.
    ///   <para>
    ///     The feature name is any fully-qualified URI.  It is
    ///     possible for an XMLReader to expose a feature value but
    ///     to be unable to change the current value.
    ///     Some feature values may be immutable or mutable only
    ///     in specific contexts, such as before, during, or after
    ///     a parse.
    ///   </para>
    ///   <para>
    ///     All XMLReaders are required to support setting
    ///     http://xml.org/sax/features/namespaces to true and
    ///     http://xml.org/sax/features/namespace-prefixes to false.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   The feature name, which is a fully-qualified URI.
    /// </param>
    /// <param name="value">
    ///   The requested value of the feature (true or false).
    /// </param>
    /// <exception cref="SAXNotRecognizedException">
    ///   If the feature value can't be assigned or retrieved.
    /// </exception>
    /// <exception cref="SAXNotSupportedException">
    ///   When the XMLReader recognizes the feature name but cannot set the requested value.
    /// </exception>
    /// <seealso cref="GetFeature" />
    void SetFeature(string name, bool value);

    /// <summary>
    ///   Look up the value of a property.
    ///   <para>
    ///     The property name is any fully-qualified URI.  It is
    ///     possible for an XMLReader to recognize a property name but
    ///     temporarily be unable to return its value.
    ///     Some property values may be available only in specific
    ///     contexts, such as before, during, or after a parse.
    ///   </para>
    ///   <para>
    ///     XMLReaders are not required to recognize any specific
    ///     property names, though an initial core set is documented for
    ///     SAX2.
    ///   </para>
    ///   <para>
    ///     Implementors are free (and encouraged) to invent their own properties,
    ///     using names built on their own URIs.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   The property name, which is a fully-qualified URI.
    /// </param>
    /// <returns>The current value of the property.</returns>
    /// <exception cref="SAXNotRecognizedException">
    ///   If the property value can't be assigned or retrieved.
    /// </exception>
    /// <exception cref="SAXNotSupportedException">
    ///   When the XMLReader recognizes the property name but cannot determine its value at this time.
    /// </exception>
    /// <seealso cref="SetProperty" />
    object GetProperty(string name);

    ////
    /// <summary>
    ///   Set the value of a property.
    ///   <para>
    ///     The property name is any fully-qualified URI.  It is
    ///     possible for an XMLReader to recognize a property name but
    ///     to be unable to change the current value.
    ///     Some property values may be immutable or mutable only
    ///     in specific contexts, such as before, during, or after
    ///     a parse.
    ///   </para>
    ///   <para>
    ///     XMLReaders are not required to recognize setting
    ///     any specific property names, though a core set is defined by
    ///     SAX2.
    ///   </para>
    ///   <para>
    ///     This method is also the standard mechanism for setting
    ///     extended handlers.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   The property name, which is a fully-qualified URI.
    /// </param>
    /// <param name="value">
    ///   The requested value for the property.
    /// </param>
    /// <exception cref="SAXNotRecognizedException">
    ///   If the property value can't be assigned or retrieved.
    /// </exception>
    /// <exception cref="SAXNotSupportedException">
    ///   When the XMLReader recognizes the property name but cannot set the requested value.
    /// </exception>
    void SetProperty(string name, object value);

    /// <summary>
    ///   Parse an XML document.
    ///   <para>
    ///     The application can use this method to instruct the XML
    ///     reader to begin parsing an XML document from any valid input
    ///     source (a character stream, a byte stream, or a URI).
    ///   </para>
    ///   <para>
    ///     Applications may not invoke this method while a parse is in
    ///     progress (they should create a new XMLReader instead for each
    ///     nested XML document).  Once a parse is complete, an
    ///     application may reuse the same XMLReader object, possibly with a
    ///     different input source.
    ///     Configuration of the XMLReader object (such as handler bindings and
    ///     values established for feature flags and properties) is unchanged
    ///     by completion of a parse, unless the definition of that aspect of
    ///     the configuration explicitly specifies other behavior.
    ///     (For example, feature flags or properties exposing
    ///     characteristics of the document being parsed.)
    ///   </para>
    ///   <para>
    ///     During the parse, the XMLReader will provide information
    ///     about the XML document through the registered event
    ///     handlers.
    ///   </para>
    ///   <para>
    ///     This method is synchronous: it will not return until parsing
    ///     has ended.  If a client application wants to terminate
    ///     parsing early, it should throw an exception.
    ///   </para>
    /// </summary>
    /// <param name="input">
    ///   The input source for the top-level of the
    ///   XML document.
    /// </param>
    /// <exception cref="SAXException">
    ///   Any SAX exception, possibly wrapping another exception.
    /// </exception>
    /// <exception cref="IOException">
    ///   An IO exception from the parser,
    ///   possibly from a byte stream or character stream
    ///   supplied by the application.
    /// </exception>
    /// <seealso cref="InputSource" />
    /// <seealso cref="Parse(string)" />
    /// <seealso cref="EntityResolver" />
    /// <seealso cref="DTDHandler" />
    /// <seealso cref="ContentHandler" />
    /// <seealso cref="ErrorHandler" />
    void Parse(InputSource input);

    /// <summary>
    ///   Parse an XML document from a system identifier (URI).
    ///   <para>
    ///     This method is a shortcut for the common case of reading a
    ///     document from a system identifier.  It is the exact
    ///     equivalent of the following:
    ///   </para>
    ///   <code>
    ///     Parse(new InputSource(systemId));
    ///   </code>
    ///   <para>
    ///     If the system identifier is a URL, it must be fully resolved
    ///     by the application before it is passed to the parser.
    ///   </para>
    /// </summary>
    /// <param name="systemId">
    ///   The system identifier (URI).
    /// </param>
    /// <exception cref="SAXException">
    ///   Any SAX exception, possibly
    ///   wrapping another exception.
    /// </exception>
    /// <exception cref="IOException">
    ///   An IO exception from the parser,
    ///   possibly from a byte stream or character stream
    ///   supplied by the application.
    /// </exception>
    /// <seealso cref="Parse(NSAX.InputSource)" />
    void Parse(string systemId);
  }
}
