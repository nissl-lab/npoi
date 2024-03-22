namespace NSAX {
  using System;
  using System.IO;
  using System.Text;

  /// <summary>
  ///   A single input source for an XML entity.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This class allows a SAX application to encapsulate information
  ///     about an input source in a single object, which may include
  ///     a public identifier, a system identifier, a byte stream (possibly
  ///     with a specified encoding), and/or a character stream.
  ///   </para>
  ///   <para>
  ///     There are two places that the application can deliver an
  ///     input source to the parser: as the argument to the Parser.parse
  ///     method, or as the return value of the EntityResolver.resolveEntity
  ///     method.
  ///   </para>
  ///   <para>
  ///     The SAX parser will use the InputSource object to determine how
  ///     to read XML input.  If there is a character stream available, the
  ///     parser will read that stream directly, disregarding any text
  ///     encoding declaration found in that stream.
  ///     If there is no character stream, but there is
  ///     a byte stream, the parser will use that byte stream, using the
  ///     encoding specified in the InputSource or else (if no encoding is
  ///     specified) autodetecting the character encoding using an algorithm
  ///     such as the one in the XML specification.  If neither a character
  ///     stream nor a
  ///     byte stream is available, the parser will attempt to open a URL
  ///     connection to the resource identified by the system
  ///     identifier.
  ///   </para>
  ///   <para>
  ///     An InputSource object belongs to the application: the SAX parser
  ///     shall never modify it in any way (it may modify a copy if
  ///     necessary).  However, standard processing of both byte and
  ///     character streams is to close them on as part of end-of-parse cleanup,
  ///     so applications should not attempt to re-use such streams after they
  ///     have been handed to a parser.
  ///   </para>
  /// </summary>
  /// <seealso cref="IXmlReader.Parse(NSAX.InputSource)" />
  /// <seealso cref="IEntityResolver.ResolveEntity" />
  /// <seealso cref="System.IO.Stream" />
  /// <seealso cref="TextReader" />
  public class InputSource {
    private TextReader _reader;
    private Stream _stream;
    private string _systemId;

    /// <summary>
    ///   Zero-argument default constructor.
    /// </summary>
    /// <seealso cref="PublicId" />
    /// <seealso cref="SystemId" />
    /// <seealso cref="Stream" />
    /// <seealso cref="Reader" />
    /// <seealso cref="Encoding" />
    public InputSource() {
    }

    /// <summary>
    ///   Create a new input source with a system identifier.
    ///   <para>
    ///     Applications may use PublicId to include a
    ///     public identifier as well, or Encoding to specify
    ///     the character encoding, if known.
    ///   </para>
    ///   <para>
    ///     If the system identifier is a URL, it must be fully
    ///     resolved (it may not be a relative URL).
    ///   </para>
    /// </summary>
    /// <param name="systemId">
    ///   The system identifier (URI).
    /// </param>
    /// <seealso cref="PublicId" />
    /// <seealso cref="SystemId" />
    /// <seealso cref="Stream" />
    /// <seealso cref="Encoding" />
    /// <seealso cref="Reader" />
    public InputSource(string systemId) {
      if (systemId == null) {
        throw new ArgumentNullException("systemId");
      }
      _systemId = systemId;
    }

    /// <summary>
    ///   Create a new input source with a byte stream.
    ///   <para>
    ///     Application writers should use SystemId to provide a base
    ///     for resolving relative URIs, may use PublicId to include a
    ///     public identifier, and may use Encoding to specify the object's
    ///     character encoding.
    ///   </para>
    /// </summary>
    /// <param name="stream">
    ///   The raw stream containing the document.
    /// </param>
    /// <seealso cref="PublicId" />
    /// <seealso cref="SystemId" />
    /// <seealso cref="Encoding" />
    /// <seealso cref="Stream" />
    /// <seealso cref="Reader" />
    public InputSource(Stream stream) {
      _stream = stream;
    }

    /// <summary>
    ///   Create a new input source with a character stream.
    ///   <para>
    ///     Application writers should use SystemId to provide a base
    ///     for resolving relative URIs, and may use PublicId to include a
    ///     public identifier.
    ///   </para>
    ///   <para>The character stream shall not include a byte order mark.</para>
    /// </summary>
    /// <seealso cref="set_PublicId" />
    /// <seealso cref="set_SystemId" />
    /// <seealso cref="set_Stream" />
    /// <seealso cref="set_Reader" />
    public InputSource(TextReader reader) {
      _reader = reader;
    }

    ////
    /// <summary>
    ///   Get or sets the public identifier for this input source.
    ///   <para>
    ///     The public identifier is always optional: if the application
    ///     writer includes one, it will be provided as part of the
    ///     location information.
    ///   </para>
    /// </summary>
    /// <returns>
    ///   The public identifier, or null if none was supplied.
    /// </returns>
    /// <seealso cref="ILocator.PublicId" />
    /// <seealso cref="SAXParseException.PublicId" />
    public virtual string PublicId { get; set; }

    /// <summary>
    ///   Get or sets the system identifier for this input source.
    ///   <para>
    ///     The Encoding property will return the character encoding
    ///     of the object pointed to, or null if unknown.
    ///   </para>
    ///   <para>If the system ID is a URL, it will be fully resolved.</para>
    ///   <para>
    ///     The system identifier is optional if there is a byte stream
    ///     or a character stream, but it is still useful to provide one,
    ///     since the application can use it to resolve relative URIs
    ///     and can include it in error messages and warnings (the parser
    ///     will attempt to open a connection to the URI only if
    ///     there is no byte stream or character stream specified).
    ///   </para>
    ///   <para>
    ///     If the application knows the character encoding of the
    ///     object pointed to by the system identifier, it can register
    ///     the encoding using the setEncoding method.
    ///   </para>
    ///   <para>
    ///     If the system identifier is a URL, it must be fully
    ///     resolved (it may not be a relative URL).
    ///   </para>
    /// </summary>
    /// <returns>The system identifier, or null if none was supplied.</returns>
    /// <seealso cref="Encoding" />
    /// <seealso cref="SystemId" />
    /// <seealso cref="ILocator.SystemId" />
    /// <seealso cref="SAXParseException.SystemId" />
    public virtual string SystemId {
      get { return _systemId; }
      set { _systemId = value; }
    }

    /// <summary>
    ///   Get or sets the byte stream for this input source.
    ///   <para>
    ///     The getEncoding method will return the character
    ///     encoding for this byte stream, or null if unknown.
    ///   </para>
    ///   <para>
    ///     The SAX parser will ignore this if there is also a character
    ///     stream specified, but it will use a byte stream in preference
    ///     to opening a URI connection itself.
    ///   </para>
    ///   <para>
    ///     If the application knows the character encoding of the
    ///     byte stream, it should set it with the setEncoding method.
    ///   </para>
    /// </summary>
    /// <returns>The byte stream, or null if none was supplied.</returns>
    /// <seealso cref="Encoding" />
    /// <seealso cref="Stream" />
    /// <seealso cref="Encoding" />
    /// <seealso cref="System.IO.Stream" />
    ////
    public virtual Stream Stream {
      get { return _stream; }
      set { _stream = value; }
    }

    /// <summary>
    ///   Get or sets the character encoding for a byte stream or URI.
    ///   This value will be ignored when the application provides a
    ///   character stream.
    ///   <para>
    ///     The encoding must be a string acceptable for an
    ///     XML encoding declaration (see section 4.3.3 of the XML 1.0
    ///     recommendation).
    ///   </para>
    ///   <para>
    ///     This method has no effect when the application provides a
    ///     character stream.
    ///   </para>
    /// </summary>
    /// <returns>The encoding, or null if none was supplied.</returns>
    /// <seealso cref="SystemId" />
    /// <seealso cref="Stream" />
    /// <seealso cref="Encoding" />
    public virtual Encoding Encoding { get; set; }

    /// <summary>
    ///   Get or sets the character stream for this input source.
    ///   Set the character stream for this input source.
    ///   <para>
    ///     If there is a character stream specified, the SAX parser
    ///     will ignore any byte stream and will not attempt to open
    ///     a URI connection to the system identifier.
    ///   </para>
    /// </summary>
    /// <returns>The character stream, or null if none was supplied.</returns>
    /// <seealso cref="Stream" />
    /// <seealso cref="TextReader" />
    public virtual TextReader Reader {
      get { return _reader; }
      set { _reader = value; }
    }
  }

  // end of InputSource.java
}
