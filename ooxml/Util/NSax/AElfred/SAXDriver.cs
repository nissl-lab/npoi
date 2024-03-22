namespace NSAX.AElfred {
  using System;
  using System.Collections;
  using System.IO;

  using NSAX;
  using Ext;
  using Helpers;

  /// <summary>
  ///   An enhanced SAX2 version of Microstar's &amp;AElig;lfred XML parser.
  ///   The enhancements primarily relate to significant improvements in
  ///   conformance to the XML specification, and SAX2 support.  Performance
  ///   has been improved.  However, the &amp;AElig;lfred proprietary APIs are
  ///   no longer public.  See the package level documentation for more
  ///   information.
  ///   <table border="1" width='100%' cellpadding='3' cellspacing='0'>
  ///     <tr bgcolor='#ccccff'>
  ///       <th>
  ///         <font size='+1'>Name</font>
  ///       </th>
  ///       <th>
  ///         <font size='+1'>Notes</font>
  ///       </th>
  ///     </tr>
  ///     <tr>
  ///       <td colspan=2>
  ///         <center>
  ///           <em>
  ///             Features ... URL prefix is
  ///             <b>http://xml.org/sax/features/</b>
  ///           </em>
  ///         </center>
  ///       </td>
  ///     </tr>
  ///     <tr>
  ///       <td>(URL)/external-general-entities</td>
  ///       <td>Value is fixed at <em>true</em></td>
  ///     </tr>
  ///     <tr>
  ///       <td>(URL)/external-parameter-entities</td>
  ///       <td>Value is fixed at <em>true</em></td>
  ///     </tr>
  ///     <tr>
  ///       <td>(URL)/namespace-prefixes</td>
  ///       <td>
  ///         Value defaults to <em>false</em> (but XML 1.0 names are
  ///         always reported)
  ///       </td>
  ///     </tr>
  ///     <tr>
  ///       <td>(URL)/namespaces</td>
  ///       <td>Value defaults to <em>true</em></td>
  ///     </tr>
  ///     <tr>
  ///       <td>(URL)/string-interning</td>
  ///       <td>Value is fixed at <em>true</em></td>
  ///     </tr>
  ///     <tr>
  ///       <td>(URL)/validation</td>
  ///       <td>Value is fixed at <em>false</em></td>
  ///     </tr>
  ///     <tr>
  ///       <td colspan=2>
  ///         <center>
  ///           <em>
  ///             Handler Properties ... URL prefix is
  ///             <b>http://xml.org/sax/properties/</b>
  ///           </em>
  ///         </center>
  ///       </td>
  ///     </tr>
  ///     <tr>
  ///       <td>(URL)/declaration-handler</td>
  ///       <td>
  ///         A declaration handler may be provided.  Declaration of general
  ///         entities is exposed, but not parameter entities; none of the entity
  ///         names reported here will begin with "%".
  ///       </td>
  ///     </tr>
  ///     <tr>
  ///       <td>(URL)/lexical-handler</td>
  ///       <td>
  ///         A lexical handler may be provided.  Entity boundaries and
  ///         comments are not exposed; only CDATA sections and the start/end of
  ///         the DTD (the internal subset is not detectible).
  ///       </td>
  ///     </tr>
  ///   </table>
  ///   <para>
  ///     Note that the declaration handler doesn't suffice for showing all
  ///     the logical structure
  ///     of the DTD; it doesn't expose the name of the root element, or the values
  ///     that are permitted in a NOTATIONS attribute.  (The former is exposed as
  ///     lexical data, and SAX2 beta doesn't expose the latter.)
  ///   </para>
  ///   <para>
  ///     Although support for several features and properties is "built in"
  ///     to this parser, it support all others by storing the assigned values
  ///     and returning them.
  ///   </para>
  ///   <para>
  ///     This parser currently implements the SAX1 Parser API, but
  ///     it may not continue to do so in the future.
  ///   </para>
  /// </summary>
  /// <seealso cref="IXmlReader" />
  public class SAXDriver : ILocator, IAttributes, IXmlReader {
    private const string FEATURE = "http://xml.org/sax/features/";
    private const string HANDLER = "http://xml.org/sax/properties/";
    private readonly ArrayList _attributeLocalNames = new ArrayList();
    private readonly ArrayList _attributeNames = new ArrayList();
    private readonly ArrayList _attributeNamespaces = new ArrayList();
    private readonly ArrayList _attributeValues = new ArrayList();
    private readonly DefaultHandler _base = new DefaultHandler();
    private readonly Stack _entityStack = new Stack();
    private readonly string[] _nsTemp = new string[3];
    private readonly NamespaceSupport _prefixStack = new NamespaceSupport();
    private int _attributeCount;
    private IContentHandler _contentHandler;
    private IDeclHandler _declHandler;
    private IDTDHandler _dtdHandler;
    private string _elementName;
    private IEntityResolver _entityResolver;
    private IErrorHandler _errorHandler;
    private Hashtable _features;
    private ILexicalHandler _lexicalHandler;
    private bool _namespaces = true;
    private bool _nspending; // indicates an attribute was read before its
    private XmlParser _parser;
    private Hashtable _properties;
    private bool _xmlNames;

    /// <summary>
    ///   Constructs a SAX Parser.
    /// </summary>
    public SAXDriver() {
      _entityResolver = _base;
      _contentHandler = _base;
      _dtdHandler = _base;
      _errorHandler = _base;
      _declHandler = _base;
      _lexicalHandler = _base;
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public virtual int Length {
      get { return _attributeNames.Count; }
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetUri(int index) {
      return (string)(_attributeNamespaces[index]);
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetLocalName(int index) {
      return (string)(_attributeLocalNames[index]);
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetQName(int i) {
      return (string)(_attributeNames[i]);
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetType(int i) {
      switch (_parser.getAttributeType(_elementName, GetQName(i))) {
        case XmlParser.ATTRIBUTE_UNDECLARED:
        case XmlParser.ATTRIBUTE_CDATA:
          return "CDATA";
        case XmlParser.ATTRIBUTE_ID:
          return "ID";
        case XmlParser.ATTRIBUTE_IDREF:
          return "IDREF";
        case XmlParser.ATTRIBUTE_IDREFS:
          return "IDREFS";
        case XmlParser.ATTRIBUTE_ENTITY:
          return "ENTITY";
        case XmlParser.ATTRIBUTE_ENTITIES:
          return "ENTITIES";
        case XmlParser.ATTRIBUTE_ENUMERATED:
          // XXX doesn't have a way to return permitted enum values,
          // though they must each be a NMTOKEN
        case XmlParser.ATTRIBUTE_NMTOKEN:
          return "NMTOKEN";
        case XmlParser.ATTRIBUTE_NMTOKENS:
          return "NMTOKENS";
        case XmlParser.ATTRIBUTE_NOTATION:
          // XXX doesn't have a way to return the permitted values,
          // each of which must be name a declared notation
          return "NOTATION";
      }

      return null;
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetValue(int i) {
      return (string)(_attributeValues[i]);
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public int GetIndex(string uri, string local) {
      int length = Length;

      for (int i = 0; i < length; i++) {
        if (!GetUri(i).Equals(uri)) {
          continue;
        }
        if (GetLocalName(i).Equals(local)) {
          return i;
        }
      }
      return -1;
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public int GetIndex(string xmlName) {
      int length = Length;

      for (int i = 0; i < length; i++) {
        if (GetQName(i).Equals(xmlName)) {
          return i;
        }
      }
      return -1;
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetType(string uri, string local) {
      int index = GetIndex(uri, local);

      if (index < 0) {
        return null;
      }
      return GetType(index);
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetType(string xmlName) {
      int index = GetIndex(xmlName);

      if (index < 0) {
        return null;
      }
      return GetType(index);
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetValue(string uri, string local) {
      int index = GetIndex(uri, local);

      if (index < 0) {
        return null;
      }
      return GetValue(index);
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetValue(string xmlName) {
      int index = GetIndex(xmlName);

      if (index < 0) {
        return null;
      }
      return GetValue(index);
    }

    /// <summary>
    ///   ILocator method (don't invoke on parser);
    /// </summary>
    public string PublicId {
      get {
        return null; // XXX track public IDs too
      }
    }

    /// <summary>
    ///   ILocator method (don't invoke on parser);
    /// </summary>
    public string SystemId {
      get { return (string)_entityStack.Peek(); }
    }

    /// <summary>
    ///   ILocator method (don't invoke on parser);
    /// </summary>
    public int LineNumber {
      get { return _parser.LineNumber; }
    }

    /// <summary>
    ///   ILocator method (don't invoke on parser);
    /// </summary>
    public int ColumnNumber {
      get { return _parser.ColumnNumber; }
    }

    /// <summary>
    ///   Returns the object used when resolving external
    ///   entities during parsing (both general and parameter entities).
    /// </summary>
    public virtual IEntityResolver EntityResolver {
      get { return _entityResolver; }
      set { _entityResolver = value ?? _base; }
    }

    /// <summary>
    ///   Returns the object used to process declarations related
    ///   to notations and unparsed entities.
    /// </summary>
    public virtual IDTDHandler DTDHandler {
      get { return _dtdHandler; }
      set { _dtdHandler = value ?? _base; }
    }

    /// <summary>
    ///   Returns the object used to report the logical
    ///   content of an XML document.
    /// </summary>
    public virtual IContentHandler ContentHandler {
      get { return _contentHandler; }
      set { _contentHandler = value ?? _base; }
    }

    /// <summary>
    ///   Returns the object used to receive callbacks for XML
    ///   errors of all levels (fatal, nonfatal, warning); this is never null;
    /// </summary>
    public virtual IErrorHandler ErrorHandler {
      get { return _errorHandler; }
      set { _errorHandler = value ?? _base; }
    }

    /// <summary>
    ///   Auxiliary API to parse an XML document, used mostly
    ///   when no URI is available.
    ///   If you want anything useful to happen, you should set
    ///   at least one type of handler.
    /// </summary>
    /// <param name="source">
    ///   The XML input source.  Don't set 'encoding' unless
    ///   you know for a fact that it's correct.
    /// </param>
    /// <seealso cref="EntityResolver" />
    /// <seealso cref="DTDHandler" />
    /// <seealso cref="ContentHandler" />
    /// <seealso cref="ErrorHandler" />
    /// <exception cref="SAXException">
    ///   The handlers may throw any SAXException,
    ///   and the parser normally throws SAXParseException objects.
    /// </exception>
    /// <exception cref="IOException">
    ///   IOExceptions are normally through through
    ///   the parser if there are problems reading the source document.
    /// </exception>
    public virtual void Parse(InputSource source) {
      lock (_base) {
        _parser = new XmlParser();
        _parser.setHandler(this);

        try {
          string systemId = source.SystemId;
          // MHK addition. SAX2 says the systemId supplied must be absolute.
          // But often it isn't. This code tries, if necessary, to expand it
          // relative to the current working directory

          systemId = TryToExpand(systemId);

          // duplicate first entry, in case startDocument handler
          // needs to use Locator.getSystemId(), before entities
          // start to get reported by the parser

          //if (systemId != null)
          _entityStack.Push(systemId);
          //else    // can't happen after tryToExpand()
          //    entityStack.push ("illegal:unknown system ID");

          _parser.DoParse(systemId, source.PublicId, source.Reader, source.Stream, source.Encoding);
        } catch (SAXException e) {
          throw;
        } catch (IOException e) {
          throw;
        } catch (Exception e) {
          throw new SAXException(e.Message, e);
        } finally {
          _contentHandler.EndDocument();
          _entityStack.Clear();
        }
      }
    }

    /// <summary>
    ///   Preferred API to parse an XML document, using a
    ///   system identifier (URI).
    /// </summary>
    public virtual void Parse(string systemId) {
      Parse(new InputSource(systemId));
    }

    /// <summary>
    ///   Tells the value of the specified feature flag.
    /// </summary>
    /// <exception cref="SAXNotRecognizedException">
    ///   thrown if the feature flag
    ///   is neither built in, nor yet assigned.
    /// </exception>
    public virtual bool GetFeature(string featureId) {
      if ((FEATURE + "validation").Equals(featureId)) {
        return false;
      }

      // external entities (both types) are currently always included
      if ((FEATURE + "external-general-entities").Equals(featureId)
          || (FEATURE + "external-parameter-entities").Equals(featureId)) {
        return true;
      }

      // element/attribute names are as written in document; no mangling
      if ((FEATURE + "namespace-prefixes").Equals(featureId)) {
        return _xmlNames;
      }

      // report element/attribute namespaces?
      if ((FEATURE + "namespaces").Equals(featureId)) {
        return _namespaces;
      }

      // XXX always provides a locator ... removed in beta

      // always interns
      if ((FEATURE + "string-interning").Equals(featureId)) {
        return true;
      }

      if (_features != null && _features.ContainsKey(featureId)) {
        return ((bool)_features[featureId]);
      }

      throw new SAXNotRecognizedException(featureId);
    }

    /// <summary>
    ///    Returns the specified property.
    /// </summary>
    /// <exception cref="SAXNotRecognizedException">
    ///   thrown if the property value
    ///   is neither built in, nor yet stored.
    /// </exception>
    public virtual object GetProperty(string propertyId) {
      if ((HANDLER + "declaration-handler").Equals(propertyId)) {
        return _declHandler;
      }

      if ((HANDLER + "lexical-handler").Equals(propertyId)) {
        return _lexicalHandler;
      }

      if (_properties != null && _properties.ContainsKey(propertyId)) {
        return _properties[propertyId];
      }

      // unknown properties
      throw new SAXNotRecognizedException(propertyId);
    }

    /// <summary>
    ///    Returns the specified property.
    /// </summary>
    /// <exception cref="SAXNotRecognizedException">
    ///   thrown if the property value
    ///   is neither built in, nor yet stored.
    /// </exception>
    public virtual void SetFeature(string featureId, bool state) {
      try {
        // Features with a defined value, we just change it if we can.
        bool value = GetFeature(featureId);

        if (state == value) {
          return;
        }

        if ((FEATURE + "namespace-prefixes").Equals(featureId)) {
          // in this implementation, this only affects xmlns reporting
          _xmlNames = state;
          return;
        }

        if ((FEATURE + "namespaces").Equals(featureId)) {
          // XXX if not currently parsing ...
          if (true) {
            _namespaces = state;
            return;
          }
          // if in mid-parse, critical info hasn't been computed/saved
        }

        // can't change builtins
        if (_features == null || !_features.ContainsKey(featureId)) {
          throw new SAXNotSupportedException(featureId);
        }
      } catch (SAXNotRecognizedException e) {
        // as-yet unknown features
        if (_features == null) {
          _features = new Hashtable(5);
        }
      }

      // record first value, or modify existing one
      _features.Add(featureId, state ? bool.TrueString : bool.FalseString);
    }

    /// <summary>
    ///    Assigns the specified property.  Like SAX1 handlers,
    ///   these may be changed at any time.
    /// </summary>
    public virtual void SetProperty(string propertyId, object property) {
      try {
        // Properties with a defined value, we just change it if we can.
        GetProperty(propertyId);

        if ((HANDLER + "declaration-handler").Equals(propertyId)) {
          if (property == null) {
            _declHandler = _base;
          } else if (! (property is IDeclHandler)) {
            throw new SAXNotSupportedException(propertyId);
          } else {
            _declHandler = (IDeclHandler)property;
          }
          return;
        }

        if ((HANDLER + "lexical-handler").Equals(propertyId)
            || "http://xml.org/sax/handlers/LexicalHandler".Equals(propertyId)) {
          // the latter name is used in some SAX2 beta software
          if (property == null) {
            _lexicalHandler = _base;
          } else if (! (property is ILexicalHandler)) {
            throw new SAXNotSupportedException(propertyId);
          } else {
            _lexicalHandler = (ILexicalHandler)property;
          }
          return;
        }

        // can't change builtins
        if (_properties == null || !_properties.ContainsKey(propertyId)) {
          throw new SAXNotSupportedException(propertyId);
        }
      } catch (SAXNotRecognizedException e) {
        // as-yet unknown properties
        if (_properties == null) {
          _properties = new Hashtable(5);
        }
      }

      // record first value, or modify existing one
      _properties.Add(propertyId, property);
    }

    public virtual void StartDocument() {
      _contentHandler.SetDocumentLocator(this);
      _contentHandler.StartDocument();
      _attributeNames.Clear();
      _attributeValues.Clear();
    }

    public virtual void EndDocument() {
      // SAX says endDocument _must_ be called (handy to close
      // files etc) so it's in a "finally" clause
    }

    public virtual object ResolveEntity(string publicId, string systemId) {
      InputSource source = _entityResolver.ResolveEntity(publicId, systemId);

      if (source == null) {
        return null;
      }
      if (source.Stream != null) {
        return source.Stream;
      }
      if (source.Stream != null) {
        if (source.Encoding == null) {
          return source.Stream;
        }
        try {
          return new StreamReader(source.Stream, source.Encoding);
        } catch (IOException e) {
          return source.Stream;
        }
      }
      string sysId = source.SystemId;
      return TryToExpand(sysId); // MHK addition
      // XXX no way to tell AElfred about new public
      // or system ids ... so relative URL resolution
      // through that entity could be less than reliable.
    }

    public virtual void StartExternalEntity(string systemId) {
      _entityStack.Push(systemId);
    }

    public virtual void EndExternalEntity(string systemId) {
      _entityStack.Pop();
    }

    public virtual void DoctypeDecl(string name, string publicId, string systemId) {
      _lexicalHandler.StartDTD(name, publicId, systemId);

      // ... the "name" is a declaration and should be given
      // to the DeclHandler (but sax2 beta doesn't).

      // the IDs for the external subset are lexical details,
      // as are the contents of the internal subset; but sax2
      // beta only provides the external subset "pre-parse"
    }

    public virtual void EndDoctype() {
      // NOTE:  some apps may care that comments and PIs,
      // are stripped from their DTD declaration context,
      // and that those declarations are themselves quite
      // thoroughly reordered here.

      DeliverDTDEvents();
      _lexicalHandler.EndDTD();
    }

    public virtual void Attribute(string aname, string value, bool isSpecified) {
      // Code changed by MHK 16 April 2001.
      // The only safe thing to do is to process all the namespace declarations
      // first, then process the ordinary attributes. So if this is a namespace
      // declaration, we deal with it now, otherwise we save it till we get the
      // startElement call.

      if (_attributeCount++ == 0) {
        if (_namespaces) {
          _prefixStack.PushContext();
        }
      }

      // set nsTemp [0] == namespace URI (or empty)
      // set nsTemp [1] == local name (or empty)
      if (value == null) {
        // MHK: I think this can only happen on an error recovery path
        // MHK: I was wrong: AElfred was notifying null values of attribute
        // declared in the DTD as #IMPLIED. But I've now changed it so it doesn't.
        return;
      }

      if (_namespaces && aname.StartsWith("xmlns")) {
        if (aname.Length == 5) {
          _prefixStack.DeclarePrefix("", value);
          //System.err.println("Declare default prefix = "+value);
          _contentHandler.StartPrefixMapping("", value);
        } else if (aname[5] == ':' && !aname.Equals("xmlns:xml")) {
          if (aname.Length == 6) {
            _errorHandler.Error(new SAXParseException("Missing namespace prefix in namespace declaration: " + aname,
              this));
            return;
          }
          string prefix = aname.Substring(6);

          if (value.Length == 0) {
            _errorHandler.Error(new SAXParseException("Missing URI in namespace declaration: " + aname, this));
            return;
          }
          _prefixStack.DeclarePrefix(prefix, value);
          //System.err.println("Declare prefix " +prefix+"="+value);
          _contentHandler.StartPrefixMapping(prefix, value);
        }

        if (!_xmlNames) {
          // if xmlNames option wasn't selected,
          // we don't report xmlns:* declarations as attributes
          return;
        }
      }

      _attributeNames.Add(aname);
      _attributeValues.Add(value);
    }

    public virtual void StartElement(string elname) {
      IContentHandler handler = _contentHandler;

      if (_attributeCount == 0) {
        _prefixStack.PushContext();
      }

      // save element name so attribute callbacks work
      _elementName = elname;
      if (_namespaces) {
        // Expand namespace prefix for all attributes
        if (_attributeCount > 0) {
          for (int i = 0; i < _attributeNames.Count; i++) {
            var aname = (string)_attributeNames[i];
            if (aname.IndexOf(':') > 0) {
              if (_xmlNames && aname.StartsWith("xmlns:")) {
                _attributeNamespaces.Add("");
                _attributeLocalNames.Add(aname);
              } else if (_prefixStack.ProcessName(aname, _nsTemp, true) == null) {
                _errorHandler.Error(new SAXParseException("undeclared name prefix in: " + aname, this));
                // recovery action: substitute a name in default namespace
                _attributeNamespaces.Add("");
                _attributeLocalNames.Add(aname.Substring(aname.IndexOf(':') + 1));
              } else {
                _attributeNamespaces.Add(_nsTemp[0]);
                _attributeLocalNames.Add(_nsTemp[1]);
              }
            } else {
              _attributeNamespaces.Add("");
              _attributeLocalNames.Add(aname);
            }
            // check uniquess of the attribute expanded name
            for (int j = 0; j < i; j++) {
              if (_attributeNamespaces[i] == _attributeNamespaces[j]
                  && _attributeLocalNames[i] == _attributeLocalNames[j]) {
                _errorHandler.Error(new SAXParseException("duplicate attribute name: " + _attributeLocalNames[j], this));
              }
            }
          }
        }

        if (_prefixStack.ProcessName(elname, _nsTemp, false) == null) {
          _errorHandler.Error(new SAXParseException("undeclared name prefix in: " + elname, this));
          _nsTemp[0] = _nsTemp[1] = "";
          // recovery action
          elname = elname.Substring(elname.IndexOf(':') + 1);
        }
        handler.StartElement(_nsTemp[0], _nsTemp[1], elname, this);
      } else {
        handler.StartElement("", "", elname, this);
      }
      // elementName = null;

      // elements with no attributes are pretty common!
      if (_attributeCount != 0) {
        _attributeNames.Clear();
        _attributeNamespaces.Clear();
        _attributeLocalNames.Clear();
        _attributeValues.Clear();
        _attributeCount = 0;
      }
      _nspending = false;
    }

    public virtual void EndElement(string elname) {
      IContentHandler handler = _contentHandler;

      if (!_namespaces) {
        handler.EndElement("", "", elname);
        return;
      }

      // following code added by MHK to fix bug Saxon 6.1/013
      if (_prefixStack.ProcessName(elname, _nsTemp, false) == null) {
        // shouldn't happen
        _errorHandler.Error(new SAXParseException("undeclared name prefix in: " + elname, this));
        _nsTemp[0] = _nsTemp[1] = "";
        elname = elname.Substring(elname.IndexOf(':') + 1);
      }

      handler.EndElement(_nsTemp[0], _nsTemp[1], elname);

      // previous code (clearly wrong): handler.endElement ("", "", elname);

      // end of MHK addition

      IEnumerator prefixes = _prefixStack.GetDeclaredPrefixes().GetEnumerator();

      while (prefixes.MoveNext()) {
        handler.EndPrefixMapping((string)prefixes.Current);
      }
      _prefixStack.PopContext();
    }

    public virtual void StartCDATA() {
      _lexicalHandler.StartCDATA();
    }

    public virtual void CharData(char[] ch, int start, int length) {
      _contentHandler.Characters(ch, start, length);
    }

    public virtual void EndCDATA() {
      _lexicalHandler.EndCDATA();
    }

    public virtual void IgnorableWhitespace(char[] ch, int start, int length) {
      _contentHandler.IgnorableWhitespace(ch, start, length);
    }

    public virtual void ProcessingInstruction(string target, string data) {
      // XXX if within DTD, perhaps it's best to discard
      // PIs since the decls to which they (usually)
      // apply get significantly rearranged

      _contentHandler.ProcessingInstruction(target, data);
    }

    public virtual void Comment(char[] ch, int start, int length) {
      // XXX if within DTD, perhaps it's best to discard
      // comments since the decls to which they (usually)
      // apply get significantly rearranged

      if (_lexicalHandler != _base) {
        _lexicalHandler.Comment(ch, start, length);
      }
    }

    // AElfred only has fatal errors
    public virtual void Error(string message, string url, int line, int column) {
      var fatal = new SAXParseException(message, null, url, line, column);
      _errorHandler.FatalError(fatal);

      // Even if the application can continue ... we can't!
      throw fatal;
    }

    //
    // Before the endDtd event, deliver all non-PE declarations.
    //
    private void DeliverDTDEvents() {
      string ename;
      string nname;
      string publicId;
      string systemId;

      // First, report all notations.
      if (_dtdHandler != _base) {
        IEnumerable notationNames = _parser.declaredNotations();

        foreach (object elem in  notationNames) {
          nname = (string)elem;
          publicId = _parser.getNotationPublicId(nname);
          systemId = _parser.getNotationSystemId(nname);
          _dtdHandler.NotationDecl(nname, publicId, systemId);
        }
      }

      // Next, report all entities.
      if (_dtdHandler != _base || _declHandler != _base) {
        IEnumerable entityNames = _parser.declaredEntities();
        int type;

        foreach (object elem in entityNames) {
          ename = (string)elem;
          type = _parser.getEntityType(ename);

          if (ename[0] == '%') {
            continue;
          }

          // unparsed
          if (type == XmlParser.ENTITY_NDATA) {
            publicId = _parser.getEntityPublicId(ename);
            systemId = _parser.getEntitySystemId(ename);
            nname = _parser.getEntityNotationName(ename);
            _dtdHandler.UnparsedEntityDecl(ename, publicId, systemId, nname);

            // external parsed
          } else if (type == XmlParser.ENTITY_TEXT) {
            publicId = _parser.getEntityPublicId(ename);
            systemId = _parser.getEntitySystemId(ename);
            _declHandler.ExternalEntityDecl(ename, publicId, systemId);

            // internal parsed
          } else if (type == XmlParser.ENTITY_INTERNAL) {
            // filter out the built-ins; even if they were
            // declared, they didn't need to be.
            if ("lt".Equals(ename) || "gt".Equals(ename) || "quot".Equals(ename) || "apos".Equals(ename)
                || "amp".Equals(ename)) {
              continue;
            }
            _declHandler.InternalEntityDecl(ename, _parser.getEntityValue(ename));
          }
        }
      }

      // elements, attributes
      if (_declHandler != _base) {
        IEnumerable elementNames = _parser.declaredElements();
        IEnumerable attNames;

        foreach (object elem in elementNames) {
          string model = null;

          ename = (string)elem;
          switch (_parser.getElementContentType(ename)) {
            case XmlParser.CONTENT_ANY:
              model = "ANY";
              break;
            case XmlParser.CONTENT_EMPTY:
              model = "EMPTY";
              break;
            case XmlParser.CONTENT_MIXED:
            case XmlParser.CONTENT_ELEMENTS:
              model = _parser.getElementContentModel(ename);
              break;
            case XmlParser.CONTENT_UNDECLARED:
            default:
              model = null;
              break;
          }
          if (model != null) {
            _declHandler.ElementDecl(ename, model);
          }

          attNames = _parser.declaredAttributes(ename);
          if (attNames != null) {
            foreach (object att in attNames) {
              var aname = (string)att;
              string type;
              string valueDefault;
              string value;

              switch (_parser.getAttributeType(ename, aname)) {
                case XmlParser.ATTRIBUTE_CDATA:
                  type = "CDATA";
                  break;
                case XmlParser.ATTRIBUTE_ENTITY:
                  type = "ENTITY";
                  break;
                case XmlParser.ATTRIBUTE_ENTITIES:
                  type = "ENTITIES";
                  break;
                case XmlParser.ATTRIBUTE_ENUMERATED:
                  type = _parser.getAttributeEnumeration(ename, aname);
                  break;
                case XmlParser.ATTRIBUTE_ID:
                  type = "ID";
                  break;
                case XmlParser.ATTRIBUTE_IDREF:
                  type = "IDREF";
                  break;
                case XmlParser.ATTRIBUTE_IDREFS:
                  type = "IDREFS";
                  break;
                case XmlParser.ATTRIBUTE_NMTOKEN:
                  type = "NMTOKEN";
                  break;
                case XmlParser.ATTRIBUTE_NMTOKENS:
                  type = "NMTOKENS";
                  break;

                  // XXX SAX2 beta doesn't have a way to return the
                  // enumerated list of permitted notations ... SAX1
                  // kluged it as NMTOKEN, but that won't work for
                  // the sort of apps that really use the DTD info
                case XmlParser.ATTRIBUTE_NOTATION:
                  type = "NOTATION";
                  break;

                default:
                  _errorHandler.FatalError(new SAXParseException("internal error, att type", this));
                  type = null;
                  break;
              }

              switch (_parser.getAttributeDefaultValueType(ename, aname)) {
                case XmlParser.ATTRIBUTE_DEFAULT_IMPLIED:
                  valueDefault = "#IMPLIED";
                  break;
                case XmlParser.ATTRIBUTE_DEFAULT_REQUIRED:
                  valueDefault = "#REQUIRED";
                  break;
                case XmlParser.ATTRIBUTE_DEFAULT_FIXED:
                  valueDefault = "#FIXED";
                  break;
                case XmlParser.ATTRIBUTE_DEFAULT_SPECIFIED:
                  valueDefault = null;
                  break;

                default:
                  _errorHandler.FatalError(new SAXParseException("internal error, att default", this));
                  valueDefault = null;
                  break;
              }

              value = _parser.getAttributeDefaultValue(ename, aname);

              _declHandler.AttributeDecl(ename, aname, type, valueDefault, value);
            }
          }
        }
      }
    }

    /// <summary>
    ///   IAttributes method (don't invoke on parser);
    /// </summary>
    public string GetName(int i) {
      return (string)(_attributeNames[i]);
    }

    public static string TryToExpand(string systemId) {
      if (systemId == null) {
        systemId = "";
      }

      if (Uri.IsWellFormedUriString(systemId, UriKind.Absolute)) {
        return systemId;
      }

      string dir;
      try {
        dir = AppDomain.CurrentDomain.BaseDirectory;
      } catch (Exception geterr) {
        // this doesn't work when running an applet
        return systemId;
      }
      if (!(dir.EndsWith("/") || systemId.StartsWith("/"))) {
        dir = dir + "/";
      }

      try {
        var currentDirectoryURL = new Uri(dir); // needs JDK 1.2
        var baseURL = new Uri(currentDirectoryURL, systemId);
        // System.err.println("SAX Driver: expanded " + systemId + " to " + baseURL);
        return baseURL.ToString();
      } catch (UriFormatException err2) {
        // go with the original one
        return systemId;
      }
    }
  }
}
