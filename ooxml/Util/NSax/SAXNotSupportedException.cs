namespace NSAX {
  using System;
  using System.Runtime.Serialization;

  /// <summary>
  ///   Exception class for an unsupported operation.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     An XMLReader will throw this exception when it recognizes a
  ///     feature or property identifier, but cannot perform the requested
  ///     operation (setting a state or value).  Other SAX2 applications and
  ///     extensions may use this class for similar purposes.
  ///   </para>
  /// </summary>
  /// <seealso cref="SAXNotRecognizedException" />
  [Serializable]
  public class SAXNotSupportedException : SAXException {
    /// <summary>
    ///   Initializes a new instance of the <see cref="T:System.Exception" /> class.
    /// </summary>
    public SAXNotSupportedException() {
    }

    /// <summary>
    ///   Initializes a new instance of the <see cref="T:System.Exception" /> class with a specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error. </param>
    public SAXNotSupportedException(string message) : base(message) {
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
    public SAXNotSupportedException(string message, Exception innerException) : base(message, innerException) {
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
    protected SAXNotSupportedException(SerializationInfo info, StreamingContext context) : base(info, context) {
    }
  }

  // end of SAXNotSupportedException.java
}
