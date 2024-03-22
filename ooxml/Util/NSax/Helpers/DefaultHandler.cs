namespace NSAX.Helpers {
  /// <summary>
  ///   Default base class for SAX2 event handlers.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This class is available as a convenience base class for SAX2
  ///     applications: it provides default implementations for all of the
  ///     callbacks in the four core SAX2 handler classes:
  ///   </para>
  ///   <ul>
  ///     <li>
  ///       <see cref="IEntityResolver" />
  ///     </li>
  ///     <li>
  ///       <see cref="IDTDHandler" />
  ///     </li>
  ///     <li>
  ///       <see cref="IContentHandler" />
  ///     </li>
  ///     <li>
  ///       <see cref="IErrorHandler" />
  ///     </li>
  ///   </ul>
  ///   <para>
  ///     Application writers can extend this class when they need to
  ///     implement only part of an interface; parser writers can
  ///     instantiate this class to provide default handlers when the
  ///     application has not supplied its own.
  ///   </para>
  /// </summary>
  /// <seealso cref="IEntityResolver" />
  /// <seealso cref="IDTDHandler" />
  /// <seealso cref="IContentHandler" />
  /// <seealso cref="IErrorHandler" />
  public class DefaultHandler : IEntityResolver, IDTDHandler, IContentHandler, IErrorHandler {
    public virtual void SetDocumentLocator(ILocator locator) {
      // no op
    }

    public virtual void StartDocument() {
      // no op
    }

    public virtual void EndDocument() {
      // no op
    }

    public virtual void StartPrefixMapping(string prefix, string uri) {
      // no op
    }

    public virtual void EndPrefixMapping(string prefix) {
      // no op
    }

    public virtual void StartElement(string uri, string localName, string qName, IAttributes atts) {
      // no op
    }

    public virtual void EndElement(string uri, string localName, string qName) {
      // no op
    }

    public virtual void Characters(char[] ch, int start, int length) {
      // no op
    }

    public virtual void IgnorableWhitespace(char[] ch, int start, int length) {
      // no op
    }

    public virtual void ProcessingInstruction(string target, string data) {
      // no op
    }

    public virtual void SkippedEntity(string name) {
      // no op
    }

    public virtual void NotationDecl(string name, string publicId, string systemId) {
      // no op
    }

    public virtual void UnparsedEntityDecl(string name, string publicId, string systemId, string notationName) {
      // no op
    }

    public virtual InputSource ResolveEntity(string publicId, string systemId) {
      return null;
    }

    public virtual void Warning(SAXParseException e) {
      // no op
    }

    public virtual void Error(SAXParseException e) {
      // no op
    }

    public virtual void FatalError(SAXParseException e) {
      throw e;
    }
  }

  // end of DefaultHandler.java
}
