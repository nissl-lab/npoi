namespace NSAX {
  using System;
  using System.Runtime.Serialization;

  /// <summary>
  ///   Encapsulate an XML parse error or warning.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This exception may include information for locating the error
  ///     in the original XML document, as if it came from a <see cref="ILocator" />object.
  ///     Note that although the application will receive a SAXParseException as the argument to the handlers
  ///     in the <see cref="IErrorHandler" /> interface,
  ///     the application is not actually required to throw the exception;
  ///     instead, it can simply read the information in it and take a
  ///     different action.
  ///   </para>
  ///   <para>
  ///     Since this exception is a subclass of <see cref="SAXException" />
  ///     it inherits the ability to wrap another exception.
  ///   </para>
  /// </summary>
  /// <seealso cref="SAXException" />
  /// <seealso cref="ILocator" />
  /// <seealso cref="IErrorHandler" />
  [Serializable]
  public class SAXParseException : SAXException {
    private int _columnNumber;
    private int _lineNumber;
    private string _publicId;
    private string _systemId;

    /// <summary>
    ///   Create a new SAXParseException from a message and a Locator.
    ///   <para>
    ///     This constructor is especially useful when an application is
    ///     creating its own exception from within a <see cref="IContentHandler" /> callback.
    ///   </para>
    /// </summary>
    /// <param name="message">
    ///   The error or warning message.
    /// </param>
    /// <param name="locator">
    ///   The locator object for the error or warning (may be null).
    /// </param>
    /// <seealso cref="ILocator" />
    public SAXParseException(string message, ILocator locator) : base(message) {
      if (locator != null) {
        Init(locator.PublicId, locator.SystemId, locator.LineNumber, locator.ColumnNumber);
      } else {
        Init(null, null, -1, -1);
      }
    }

    /// <summary>
    ///   Wrap an existing exception in a SAXParseException.
    ///   <para>
    ///     This constructor is especially useful when an application is
    ///     creating its own exception from within a <see cref="org.xml.sax.ContentHandler" /> callback,
    ///     and needs to wrap an existing exception that is not a
    ///     subclass of <see cref="SAXException" />.
    ///   </para>
    /// </summary>
    /// <param name="message">
    ///   The error or warning message, or null to
    ///   use the message from the embedded exception.
    /// </param>
    /// <param name="locator">
    ///   The locator object for the error or warning (may be null).
    /// </param>
    /// <param name="ex">Any exception.</param>
    /// <seealso cref="ILocator" />
    public SAXParseException(string message, ILocator locator, Exception ex) : base(message, ex) {
      if (locator != null) {
        Init(locator.PublicId, locator.SystemId, locator.LineNumber, locator.ColumnNumber);
      } else {
        Init(null, null, -1, -1);
      }
    }

    /// <summary>
    ///   Create a new SAXParseException.
    ///   <para>This constructor is most useful for parser writers.</para>
    ///   <para>
    ///     All parameters except the message are as if
    ///     they were provided by a <see cref="ILocator" />.  For example, if the
    ///     system identifier is a URL (including relative filename), the
    ///     caller must resolve it fully before creating the exception.
    ///   </para>
    /// </summary>
    /// <param name="message">
    ///   The error or warning message.
    /// </param>
    /// <param name="publicId">
    ///   The public identifier of the entity that generated the error or warning.
    /// </param>
    /// <param name="systemId">
    ///   The system identifier of the entity that generated the error or warning.
    /// </param>
    /// <param name="lineNumber">
    ///   The line number of the end of the text that caused the error or warning.
    /// </param>
    /// <param name="columnNumber">
    ///   The column number of the end of the text that cause the error or warning.
    /// </param>
    public SAXParseException(string message, string publicId, string systemId, int lineNumber, int columnNumber)
      : base(message) {
      Init(publicId, systemId, lineNumber, columnNumber);
    }

    /// <summary>
    ///   Create a new SAXParseException with an embedded exception.
    ///   <para>
    ///     This constructor is most useful for parser writers who
    ///     need to wrap an exception that is not a subclass of
    ///     <see cref="SAXException" />.
    ///   </para>
    ///   <para>
    ///     All parameters except the message and exception are as if
    ///     they were provided by a <see cref="ILocator" />.  For example, if the
    ///     system identifier is a URL (including relative filename), the
    ///     caller must resolve it fully before creating the exception.
    ///   </para>
    /// </summary>
    /// <param name="message">The error or warning message, or null to use the message from the embedded exception.</param>
    /// <param name="publicId">The public identifier of the entity that generated the error or warning.</param>
    /// <param name="systemId">The system identifier of the entity that generated the error or warning.</param>
    /// <param name="lineNumber">The line number of the end of the text that caused the error or warning.</param>
    /// <param name="columnNumber">The column number of the end of the text that cause the error or warning.</param>
    /// <param name="ex">Another exception to embed in this one.</param>
    public SAXParseException(string message, string publicId, string systemId, int lineNumber, int columnNumber,
                             Exception ex) : base(message, ex) {
      Init(publicId, systemId, lineNumber, columnNumber);
    }

    /// <summary>
    ///   Initializes a new instance of the <see cref="T:System.Exception" /> class with serialized data.
    /// </summary>
    /// <param name="info">
    ///   The <see cref="T:System.Runtime.Serialization.SerializationInfo" /> that holds the serialized object
    ///   data about the exception being thrown.
    /// </param>
    /// <param name="context">
    ///   The <see cref="T:System.Runtime.Serialization.StreamingContext" /> that contains contextual
    ///   information about the source or destination.
    /// </param>
    /// <exception cref="T:System.ArgumentNullException">The <paramref name="info" /> parameter is null. </exception>
    /// <exception cref="T:System.Runtime.Serialization.SerializationException">
    ///   The class name is null or
    ///   <see cref="P:System.Exception.HResult" /> is zero (0).
    /// </exception>
    protected SAXParseException(SerializationInfo info, StreamingContext context) : base(info, context) {
      _publicId = info.GetString("PublicId");
      _systemId = info.GetString("SystemId");
      _lineNumber = info.GetInt32("LineNumber");
      _columnNumber = info.GetInt32("ColumnNumber");
    }

    /// <summary>
    ///   Get the public identifier of the entity where the exception occurred.
    /// </summary>
    /// ///
    /// <returns>A string containing the public identifier, or null if none is available.</returns>
    /// <seealso cref="ILocator.PublicId" />
    public string PublicId {
      get { return _publicId; }
    }

    /// <summary>
    ///   Get the system identifier of the entity where the exception occurred.
    ///   <para>
    ///     If the system identifier is a URL, it will have been resolved
    ///     fully.
    ///   </para>
    /// </summary>
    /// <returns> A string containing the system identifier, or null if none is available.</returns>
    /// <seealso cref="ILocator.SystemId" />
    public string SystemId {
      get { return _systemId; }
    }

    /// <summary>
    ///   The line number of the end of the text where the exception occurred.
    ///   <para>The first line is line 1.</para>
    /// </summary>
    /// <returns>An integer representing the line number, or -1 if none is available.</returns>
    /// <seealso cref="ILocator.LineNumber" />
    public int LineNumber {
      get { return _lineNumber; }
    }

    /// <summary>
    ///   The column number of the end of the text where the exception occurred.
    ///   <para>The first column in a line is position 1.</para>
    /// </summary>
    /// <returns>An integer representing the column number, or -1 if none is available.</returns>
    /// <seealso cref="ILocator.ColumnNumber" />
    public int ColumnNumber {
      get { return _columnNumber; }
    }

    public override void GetObjectData(SerializationInfo info, StreamingContext context) {
      base.GetObjectData(info, context);
      info.AddValue("PublicId", PublicId);
      info.AddValue("SystemId", SystemId);
      info.AddValue("LineNumber", LineNumber);
      info.AddValue("ColumnNumber", ColumnNumber);
    }

    /// <summary>
    ///   Internal initialization method.
    /// </summary>
    /// <param name="publicId">The public identifier of the entity which generated the exception, or null.</param>
    /// <param name="systemId">The system identifier of the entity which generated the exception, or null.</param>
    /// <param name="lineNumber">The line number of the error, or -1.</param>
    /// <param name="columnNumber">The column number of the error, or -1.</param>
    private void Init(string publicId, string systemId, int lineNumber, int columnNumber) {
      _publicId = publicId;
      _systemId = systemId;
      _lineNumber = lineNumber;
      _columnNumber = columnNumber;
    }
  }

  // end of SAXParseException.java
}
