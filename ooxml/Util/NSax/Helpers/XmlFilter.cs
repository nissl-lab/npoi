namespace NSAX.Helpers {
  using System;
  using System.IO;

  /// <summary>
  ///   Base class for deriving an XML filter.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This class is designed to sit between an <see cref="IXmlReader" />
  ///     and the client application's event handlers.  By default, it
  ///     does nothing but pass requests up to the reader and events
  ///     on to the handlers unmodified, but subclasses can override
  ///     specific methods to modify the event stream or the configuration
  ///     requests as they pass through.
  ///   </para>
  /// </summary>
  /// <seealso cref="IXmlFilter" />
  /// <seealso cref="IXmlReader" />
  /// <seealso cref="IEntityResolver" />
  /// <seealso cref="IDTDHandler" />
  /// <seealso cref="IContentHandler" />
  /// <seealso cref="IErrorHandler" />
  public class XmlFilter : IXmlFilter, IEntityResolver, IDTDHandler, IContentHandler, IErrorHandler {
    private IContentHandler _contentHandler;
    private IDTDHandler _dtdHandler;
    private IEntityResolver _entityResolver;
    private IErrorHandler _errorHandler;
    private ILocator _locator;
    private IXmlReader _parent;

    /// <summary>
    ///   Construct an empty XML filter, with no parent.
    ///   <para>
    ///     This filter will have no parent: you must assign a parent
    ///     before you start a parse or do any configuration with
    ///     setFeature or setProperty, unless you use this as a pure event
    ///     consumer rather than as an <see cref="IXmlReader" />.
    ///   </para>
    /// </summary>
    /// <seealso cref="IXmlReader.SetFeature" />
    /// <seealso cref="IXmlReader.SetProperty" />
    /// <seealso cref="Parent" />
    public XmlFilter() {
    }

    /// <summary>
    ///   Construct an XML filter with the specified parent.
    /// </summary>
    /// <param name="parent">The parent</param>
    /// <seealso cref="Parent" />
    public XmlFilter(IXmlReader parent) {
      _parent = parent;
    }

    public virtual void SetDocumentLocator(ILocator locator) {
      _locator = locator;
      if (_contentHandler != null) {
        _contentHandler.SetDocumentLocator(locator);
      }
    }

    public virtual void StartDocument() {
      if (_contentHandler != null) {
        _contentHandler.StartDocument();
      }
    }

    public virtual void EndDocument() {
      if (_contentHandler != null) {
        _contentHandler.EndDocument();
      }
    }

    public virtual void StartPrefixMapping(string prefix, string uri) {
      if (_contentHandler != null) {
        _contentHandler.StartPrefixMapping(prefix, uri);
      }
    }

    public virtual void EndPrefixMapping(string prefix) {
      if (_contentHandler != null) {
        _contentHandler.EndPrefixMapping(prefix);
      }
    }

    public virtual void StartElement(string uri, string localName, string qName, IAttributes atts) {
      if (_contentHandler != null) {
        _contentHandler.StartElement(uri, localName, qName, atts);
      }
    }

    public virtual void EndElement(string uri, string localName, string qName) {
      if (_contentHandler != null) {
        _contentHandler.EndElement(uri, localName, qName);
      }
    }

    public virtual void Characters(char[] ch, int start, int length) {
      if (_contentHandler != null) {
        _contentHandler.Characters(ch, start, length);
      }
    }

    public virtual void IgnorableWhitespace(char[] ch, int start, int length) {
      if (_contentHandler != null) {
        _contentHandler.IgnorableWhitespace(ch, start, length);
      }
    }

    public virtual void ProcessingInstruction(string target, string data) {
      if (_contentHandler != null) {
        _contentHandler.ProcessingInstruction(target, data);
      }
    }

    public virtual void SkippedEntity(string name) {
      if (_contentHandler != null) {
        _contentHandler.SkippedEntity(name);
      }
    }

    public virtual void NotationDecl(string name, string publicId, string systemId) {
      if (_dtdHandler != null) {
        _dtdHandler.NotationDecl(name, publicId, systemId);
      }
    }

    public virtual void UnparsedEntityDecl(string name, string publicId, string systemId, string notationName) {
      if (_dtdHandler != null) {
        _dtdHandler.UnparsedEntityDecl(name, publicId, systemId, notationName);
      }
    }

    public virtual InputSource ResolveEntity(string publicId, string systemId) {
      if (_entityResolver != null) {
        return _entityResolver.ResolveEntity(publicId, systemId);
      }
      return null;
    }

    public virtual void Warning(SAXParseException e) {
      if (_errorHandler != null) {
        _errorHandler.Warning(e);
      }
    }

    public virtual void Error(SAXParseException e) {
      if (_errorHandler != null) {
        _errorHandler.Error(e);
      }
    }

    public virtual void FatalError(SAXParseException e) {
      if (_errorHandler != null) {
        _errorHandler.FatalError(e);
      }
    }

    public IXmlReader Parent {
      get { return _parent; }
      set { _parent = value; }
    }

    public virtual void SetFeature(string name, bool value) {
      if (_parent != null) {
        _parent.SetFeature(name, value);
      } else {
        throw new SAXNotRecognizedException("Feature: " + name);
      }
    }

    public virtual bool GetFeature(string name) {
      if (_parent != null) {
        return _parent.GetFeature(name);
      }
      throw new SAXNotRecognizedException("Feature: " + name);
    }

    public virtual void SetProperty(string name, object value) {
      if (_parent != null) {
        _parent.SetProperty(name, value);
      } else {
        throw new SAXNotRecognizedException("Property: " + name);
      }
    }

    public virtual object GetProperty(string name) {
      if (_parent != null) {
        return _parent.GetProperty(name);
      }
      throw new SAXNotRecognizedException("Property: " + name);
    }

    public IEntityResolver EntityResolver {
      get { return _entityResolver; }
      set { _entityResolver = value; }
    }

    public IDTDHandler DTDHandler {
      get { return _dtdHandler; }
      set { _dtdHandler = value; }
    }

    public IContentHandler ContentHandler {
      get { return _contentHandler; }
      set { _contentHandler = value; }
    }

    public IErrorHandler ErrorHandler {
      get { return _errorHandler; }
      set { _errorHandler = value; }
    }

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
    ///   Any SAX exception, possibly
    ///   wrapping another exception.
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
    public virtual void Parse(InputSource input) {
      SetupParse();
      _parent.Parse(input);
    }

    ////
    /// <summary>
    ///   Parse a document.
    /// </summary>
    /// <param name="systemId">
    ///   The system identifier as a fully-qualified URI.
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
    public virtual void Parse(string systemId) {
      Parse(new InputSource(systemId));
    }

    /// <summary>
    ///   Set up before a parse.
    ///   <para>
    ///     Before every parse, check whether the parent is
    ///     non-null, and re-register the filter for all of the
    ///     events.
    ///   </para>
    /// </summary>
    private void SetupParse() {
      if (_parent == null) {
        throw new InvalidOperationException("No parent for filter");
      }
      _parent.EntityResolver = this;
      _parent.DTDHandler = this;
      _parent.ContentHandler = this;
      _parent.ErrorHandler = this;
    }

    
  }

  // end of XMLFilterImpl.java
}
