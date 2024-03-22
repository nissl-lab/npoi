namespace NSAX {
  using System;
  using System.Runtime.Serialization;

  /// <summary>
  ///   Encapsulate a general SAX error or warning.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This class can contain basic error or warning information from
  ///     either the XML parser or the application: a parser writer or
  ///     application writer can subclass it to provide additional
  ///     functionality.  SAX handlers may throw this exception or
  ///     any exception subclassed from it.
  ///   </para>
  ///   <para>
  ///     If the application needs to pass through other types of
  ///     exceptions, it must wrap those exceptions in a SAXException
  ///     or an exception derived from a SAXException.
  ///   </para>
  ///   <para>
  ///     If the parser or application needs to include information about a
  ///     specific location in an XML document, it should use the
  ///     <see cref="SAXParseException" /> subclass.
  ///   </para>
  /// </summary>
  /// <seealso cref="SAXParseException" />
  [Serializable]
  public class SAXException : Exception {
    /// <summary>
    ///   Initializes a new instance of the <see cref="T:System.Exception" /> class.
    /// </summary>
    public SAXException() {
    }

    /// <summary>
    ///   Initializes a new instance of the <see cref="T:System.Exception" /> class with a specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error. </param>
    public SAXException(string message) : base(message) {
    }

    /// <summary>
    ///   Initializes a new instance of the <see cref="T:System.Exception" /> class with a specified error message and a
    ///   reference to the inner exception that is the cause of this exception.
    /// </summary>
    /// <param name="message">The error message that explains the reason for the exception. </param>
    /// <param name="innerException">
    ///   The exception that is the cause of the current exception, or a null reference (Nothing in
    ///   Visual Basic) if no inner exception is specified.
    /// </param>
    public SAXException(string message, Exception innerException) : base(message, innerException) {
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
    protected SAXException(SerializationInfo info, StreamingContext context) : base(info, context) {
    }
  }

  // end of SAXException.java
}
