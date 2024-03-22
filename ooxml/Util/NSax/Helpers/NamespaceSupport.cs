namespace NSAX.Helpers {
  using System;
  using System.Collections;
  using System.Linq;

  /// <summary>
  ///   Encapsulate Namespace logic for use by applications using SAX,
  ///   or internally by SAX drivers.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This class encapsulates the logic of Namespace processing: it
  ///     tracks the declarations currently in force for each context and
  ///     automatically processes qualified XML names into their Namespace
  ///     parts; it can also be used in reverse for generating XML qnames
  ///     from Namespaces.
  ///   </para>
  ///   <para>
  ///     Namespace support objects are reusable, but the reset method
  ///     must be invoked between each session.
  ///   </para>
  ///   <para>Here is a simple session:</para>
  ///   <code>
  ///     string parts[] = new string[3];
  ///     NamespaceSupport support = new NamespaceSupport();
  ///     support.PushContext();
  ///     support.DeclarePrefix("", "http://www.w3.org/1999/xhtml");
  ///     support.DeclarePrefix("dc", "http://www.purl.org/dc#");
  ///     parts = support.ProcessName("p", parts, false);
  ///     Console.WriteLine("Namespace URI: " + parts[0]);
  ///     Console.WriteLine("Local name: " + parts[1]);
  ///     Console.WriteLine("Raw name: " + parts[2]);
  ///     parts = support.ProcessName("dc:title", parts, false);
  ///     Console.WriteLine("Namespace URI: " + parts[0]);
  ///     Console.WriteLine("Local name: " + parts[1]);
  ///     Console.WriteLine("Raw name: " + parts[2]);
  ///     support.PopContext();
  ///   </code>
  ///   <para>
  ///     Note that this class is optimized for the use case where most
  ///     elements do not contain Namespace declarations: if the same
  ///     prefix/URI mapping is repeated for each context (for example), this
  ///     class will be somewhat less efficient.
  ///   </para>
  ///   <para>
  ///     Although SAX drivers (parsers) may choose to use this class to
  ///     implement namespace handling, they are not required to do so.
  ///     Applications must track namespace information themselves if they
  ///     want to use namespace information.
  ///   </para>
  /// </summary>
  public class NamespaceSupport {
    /// <summary>
    ///   The XML Namespace URI as a constant.
    ///   The value is <c>http://www.w3.org/XML/1998/namespace</c>
    ///   as defined in the "Namespaces in XML" * recommendation.
    ///   <para>This is the Namespace URI that is automatically mapped to the "xml" prefix.</para>
    /// </summary>
    public const string XMLNS = "http://www.w3.org/XML/1998/namespace";

    /// <summary>
    ///   The namespace declaration URI as a constant.
    ///   The value is <c>http://www.w3.org/xmlns/2000/</c>, as defined
    ///   in a backwards-incompatible erratum to the "Namespaces in XML"
    ///   recommendation.  Because that erratum postdated SAX2, SAX2 defaults
    ///   to the original recommendation, and does not normally use this URI.
    ///   <para>
    ///     This is the Namespace URI that is optionally applied to
    ///     <em>xmlns</em> and <em>xmlns:*</em> attributes, which are used to
    ///     declare namespaces.
    ///   </para>
    /// </summary>
    /// <seealso cref="SetNamespaceDeclUris" />
    /// <seealso cref="IsNamespaceDeclUris" />
    ////
    public const string NSDECL = "http://www.w3.org/xmlns/2000/";
    internal bool NamespaceDeclUris;
    private int _contextPos;
    private Context[] _contexts;
    private Context _currentContext;

    /// <summary>
    ///   Create a new Namespace support object.
    /// </summary>
    public NamespaceSupport() {
      Reset();
    }

    /// <summary>
    ///   Returns true if namespace declaration attributes are placed into
    ///   a namespace.  This behavior is not the default.
    /// </summary>
    public bool IsNamespaceDeclUris {
      get { return NamespaceDeclUris; }
    }

    /// <summary>
    ///   Reset this Namespace support object for reuse.
    ///   <para>
    ///     It is necessary to invoke this method before reusing the
    ///     Namespace support object for a new session.  If namespace
    ///     declaration URIs are to be supported, that flag must also
    ///     be set to a non-default value.
    ///   </para>
    /// </summary>
    /// <seealso cref="SetNamespaceDeclUris" />
    public void Reset() {
      _contexts = new Context[32];
      NamespaceDeclUris = false;
      _contextPos = 0;
      _contexts[_contextPos] = _currentContext = new Context();
      _currentContext.DeclarePrefix("xml", XMLNS);
    }

    /// <summary>
    ///   Start a new Namespace context.
    ///   The new context will automatically inherit
    ///   the declarations of its parent context, but it will also keep
    ///   track of which declarations were made within this context.
    ///   <para>
    ///     Event callback code should start a new context once per element.
    ///     This means being ready to call this in either of two places.
    ///     For elements that don't include namespace declarations, the
    ///     <em>ContentHandler.startElement()</em> callback is the right place.
    ///     For elements with such a declaration, it'd done in the first
    ///     <em>ContentHandler.startPrefixMapping()</em> callback.
    ///     A boolean flag can be used to
    ///     track whether a context has been started yet.  When either of
    ///     those methods is called, it checks the flag to see if a new context
    ///     needs to be started.  If so, it starts the context and sets the
    ///     flag.  After <em>ContentHandler.startElement()</em>
    ///     does that, it always clears the flag.
    ///   </para>
    ///   <para>
    ///     Normally, SAX drivers would push a new context at the beginning
    ///     of each XML element.  Then they perform a first pass over the
    ///     attributes to process all namespace declarations, making
    ///     <em>ContentHandler.startPrefixMapping()</em> callbacks.
    ///     Then a second pass is made, to determine the namespace-qualified
    ///     names for all attributes and for the element name.
    ///     Finally all the information for the
    ///     <em>ContentHandler.startElement()</em> callback is available,
    ///     so it can then be made.
    ///   </para>
    ///   <para>
    ///     The Namespace support object always starts with a base context
    ///     already in force: in this context, only the "xml" prefix is
    ///     declared.
    ///   </para>
    /// </summary>
    /// <seealso cref="IContentHandler" />
    /// <seealso cref="PopContext" />
    public void PushContext() {
      int max = _contexts.Length;

      _contexts[_contextPos].DeclsOK = false;
      _contextPos++;

      // Extend the array if necessary
      if (_contextPos >= max) {
        var newContexts = new Context[max * 2];
        Array.Copy(_contexts, 0, newContexts, 0, max);
        max *= 2;
        _contexts = newContexts;
      }

      // Allocate the context if necessary.
      _currentContext = _contexts[_contextPos];
      if (_currentContext == null) {
        _contexts[_contextPos] = _currentContext = new Context();
      }

      // Set the parent, if any.
      if (_contextPos > 0) {
        _currentContext.SetParent(_contexts[_contextPos - 1]);
      }
    }

    /// <summary>
    ///   Revert to the previous Namespace context.
    ///   <para>
    ///     Normally, you should pop the context at the end of each
    ///     XML element.  After popping the context, all Namespace prefix
    ///     mappings that were previously in force are restored.
    ///   </para>
    ///   <para>
    ///     You must not attempt to declare additional Namespace
    ///     prefixes after popping a context, unless you push another
    ///     context first.
    ///   </para>
    /// </summary>
    /// <seealso cref="PushContext" />
    public void PopContext() {
      _contexts[_contextPos].Clear();
      _contextPos--;
      if (_contextPos < 0) {
        throw new InvalidOperationException();
      }
      _currentContext = _contexts[_contextPos];
    }

    /// <summary>
    ///   Declare a Namespace prefix.  All prefixes must be declared
    ///   before they are referenced.  For example, a SAX driver (parser)
    ///   would scan an element's attributes
    ///   in two passes:  first for namespace declarations,
    ///   then a second pass using <see cref="ProcessName" /> to
    ///   interpret prefixes against (potentially redefined) prefixes.
    ///   <para>
    ///     This method declares a prefix in the current Namespace
    ///     context; the prefix will remain in force until this context
    ///     is popped, unless it is shadowed in a descendant context.
    ///   </para>
    ///   <para>
    ///     To declare the default element Namespace, use the empty string as
    ///     the prefix.
    ///   </para>
    ///   <para>
    ///     Note that you must <em>not</em> declare a prefix after
    ///     you've pushed and popped another Namespace context, or
    ///     treated the declarations phase as complete by processing
    ///     a prefixed name.
    ///   </para>
    ///   <para>
    ///     Note that there is an asymmetry in this library: <see cref="GetPrefix" /> will not return the "" prefix,
    ///     even if you have declared a default element namespace.
    ///     To check for a default namespace,
    ///     you have to look it up explicitly using <see cref="GetUri" />.
    ///     This asymmetry exists to make it easier to look up prefixes
    ///     for attribute names, where the default prefix is not allowed.
    ///   </para>
    /// </summary>
    /// <param name="prefix">
    ///   The prefix to declare, or the empty string to
    ///   indicate the default element namespace.  This may never have
    ///   the value "xml" or "xmlns".
    /// </param>
    /// <param name="uri">
    ///   The Namespace URI to associate with the prefix.
    /// </param>
    /// <returns><c>true</c> if the prefix was legal, <c>false</c> otherwise</returns>
    /// <seealso cref="ProcessName" />
    /// <seealso cref="GetUri" />
    /// <seealso cref="GetPrefix" />
    public bool DeclarePrefix(string prefix, string uri) {
      if (prefix.Equals("xml") || prefix.Equals("xmlns")) {
        return false;
      }
      _currentContext.DeclarePrefix(prefix, uri);
      return true;
    }

    /// <summary>
    ///   Process a raw XML qualified name, after all declarations in the
    ///   current context have been handled by <see cref="DeclarePrefix" />.
    ///   <para>
    ///     This method processes a raw XML qualified name in the
    ///     current context by removing the prefix and looking it up among
    ///     the prefixes currently declared.  The return value will be the
    ///     array supplied by the caller, filled in as follows:
    ///   </para>
    ///   <dl>
    ///     <dt>parts[0]</dt>
    ///     <dd>
    ///       The Namespace URI, or an empty string if none is
    ///       in use.
    ///     </dd>
    ///     <dt>parts[1]</dt>
    ///     <dd>The local name (without prefix).</dd>
    ///     <dt>parts[2]</dt>
    ///     <dd>The original raw name.</dd>
    ///   </dl>
    ///   <para>
    ///     All of the strings in the array will be internalized.  If
    ///     the raw name has a prefix that has not been declared, then
    ///     the return value will be null.
    ///   </para>
    ///   <para>
    ///     Note that attribute names are processed differently than
    ///     element names: an unprefixed element name will receive the
    ///     default Namespace (if any), while an unprefixed attribute name
    ///     will not.
    ///   </para>
    /// </summary>
    /// <param name="qName">
    ///   The XML qualified name to be processed.
    /// </param>
    /// <param name="parts">
    ///   An array supplied by the caller, capable of
    ///   holding at least three members.
    /// </param>
    /// <param name="isAttribute">
    ///   A flag indicating whether this is an
    ///   attribute name (true) or an element name (false).
    /// </param>
    /// <returns>
    ///   The supplied array holding three internalized strings
    ///   representing the Namespace URI (or empty string), the
    ///   local name, and the XML qualified name; or null if there
    ///   is an undeclared prefix.
    /// </returns>
    /// <seealso cref="DeclarePrefix" />
    /// <seealso cref="string.Intern" />
    public string[] ProcessName(string qName, string[] parts, bool isAttribute) {
      string[] myParts = _currentContext.ProcessName(qName, isAttribute, NamespaceDeclUris);
      if (myParts == null) {
        return null;
      }
      parts[0] = myParts[0];
      parts[1] = myParts[1];
      parts[2] = myParts[2];
      return parts;
    }

    /// <summary>
    ///   Look up a prefix and get the currently-mapped Namespace URI.
    ///   <para>
    ///     This method looks up the prefix in the current context.
    ///     Use the empty string ("") for the default Namespace.
    ///   </para>
    /// </summary>
    /// <param name="prefix">
    ///   The prefix to look up.
    /// </param>
    /// <returns>
    ///   The associated Namespace URI, or null if the prefix
    ///   is undeclared in this context.
    /// </returns>
    /// <seealso cref="GetPrefix" />
    /// <seealso cref="GetPrefixes()" />
    public string GetUri(string prefix) {
      return _currentContext.GetURI(prefix);
    }

    /// <summary>
    ///   Return an enumeration of all prefixes whose declarations are
    ///   active in the current context.
    ///   This includes declarations from parent contexts that have
    ///   not been overridden.
    ///   <para>
    ///     <strong>Note:</strong> if there is a default prefix, it will not be
    ///     returned in this enumeration; check for the default prefix
    ///     using the <see cref="GetUri" /> with an argument of "".
    ///   </para>
    /// </summary>
    /// <returns>An enumeration of prefixes (never empty).</returns>
    /// <seealso cref="GetDeclaredPrefixes" />
    /// <seealso cref="GetUri" />
    public IEnumerable GetPrefixes() {
      return _currentContext.GetPrefixes();
    }

    /// <summary>
    ///   Return one of the prefixes mapped to a Namespace URI.
    ///   <para>
    ///     If more than one prefix is currently mapped to the same
    ///     URI, this method will make an arbitrary selection; if you
    ///     want all of the prefixes, use the <see cref="GetPrefixes()" />
    ///     method instead.
    ///   </para>
    ///   <para>
    ///     <strong>Note:</strong> this will never return the empty (default) prefix;
    ///     to check for a default prefix, use the <see cref="GetUri" />
    ///     method with an argument of "".
    ///   </para>
    /// </summary>
    /// <param name="uri">
    ///   the namespace URI
    /// </param>
    /// <returns>
    ///   one of the prefixes currently mapped to the URI supplied,
    ///   or null if none is mapped or if the URI is assigned to
    ///   the default namespace
    /// </returns>
    /// <seealso cref="GetPrefixes(string)" />
    /// <seealso cref="GetUri" />
    public string GetPrefix(string uri) {
      return _currentContext.GetPrefix(uri);
    }

    /// <summary>
    ///   Return an enumeration of all prefixes for a given URI whose
    ///   declarations are active in the current context.
    ///   This includes declarations from parent contexts that have
    ///   not been overridden.
    ///   <para>
    ///     This method returns prefixes mapped to a specific Namespace
    ///     URI.  The xml: prefix will be included.  If you want only one
    ///     prefix that's mapped to the Namespace URI, and you don't care
    ///     which one you get, use the <see cref="GetPrefix" />
    ///     method instead.
    ///   </para>
    ///   <para>
    ///     <strong>Note:</strong> the empty (default) prefix is <em>never</em> included
    ///     in this enumeration; to check for the presence of a default
    ///     Namespace, use the <see cref="GetUri" /> method with an
    ///     argument of "".
    ///   </para>
    /// </summary>
    /// <param name="uri">
    ///   The Namespace URI.
    /// </param>
    /// <returns>An enumeration of prefixes (never empty).</returns>
    /// <seealso cref="GetPrefix" />
    /// <seealso cref="GetDeclaredPrefixes" />
    /// <seealso cref="GetUri" />
    public IEnumerable GetPrefixes(string uri) {
      var prefixes = new ArrayList();
      IEnumerator allPrefixes = GetPrefixes().GetEnumerator();
      while (allPrefixes.MoveNext()) {
        var prefix = (string)allPrefixes.Current;
        if (uri.Equals(GetUri(prefix))) {
          prefixes.Add(prefix);
        }
      }
      return prefixes;
    }

    /// <summary>
    ///   Return an enumeration of all prefixes declared in this context.
    ///   <para>
    ///     The empty (default) prefix will be included in this
    ///     enumeration; note that this behaviour differs from that of
    ///     <see cref="GetPrefix" /> and <see cref="GetPrefixes()" />.
    ///   </para>
    /// </summary>
    /// <returns>
    ///   An enumeration of all prefixes declared in this
    ///   context.
    /// </returns>
    /// <seealso cref="GetPrefixes()" />
    /// <seealso cref="GetUri" />
    public IEnumerable GetDeclaredPrefixes() {
      return _currentContext.GetDeclaredPrefixes();
    }

    /// <summary>
    ///   Controls whether namespace declaration attributes are placed
    ///   into the <see cref="NSDECL" /> namespace
    ///   by <see cref="ProcessName" />.  This may only be
    ///   changed before any contexts have been pushed.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    ///   when attempting to set this
    ///   after any context has been pushed.
    /// </exception>
    public void SetNamespaceDeclUris(bool value) {
      if (_contextPos != 0) {
        throw new InvalidOperationException();
      }
      if (value == NamespaceDeclUris) {
        return;
      }
      NamespaceDeclUris = value;
      if (value) {
        _currentContext.DeclarePrefix("xmlns", NSDECL);
      } else {
        _contexts[_contextPos] = _currentContext = new Context();
        _currentContext.DeclarePrefix("xml", XMLNS);
      }
    }

    /**
     * Internal class for a single Namespace context.
     *
     * <p>This module caches and reuses Namespace contexts,
     * so the number allocated
     * will be equal to the element depth of the document, not to the total
     * number of elements (i.e. 5-10 rather than tens of thousands).
     * Also, data structures used to represent contexts are shared when
     * possible (child contexts without declarations) to further reduce
     * the amount of memory that's consumed.
     * </p>
     */

    internal sealed class Context {
      /**
	 * Create the root-level Namespace context.
	 */

      internal bool DeclsOK = true;
      private Hashtable _attributeNameTable;
      private bool _declSeen;
      private ArrayList _declarations;
      private string _defaultNs;
      private Hashtable _elementNameTable;
      private Context _parent;
      private Hashtable _prefixTable;
      private Hashtable _uriTable;

      public Context() {
        CopyTables();
      }

      /**
	 * (Re)set the parent of this Namespace context.
	 * The context must either have been freshly constructed,
	 * or must have been cleared.
	 *
	 * @param context The parent Namespace context object.
	 */

      public void SetParent(Context parent) {
        _parent = parent;
        _declarations = null;
        _prefixTable = parent._prefixTable;
        _uriTable = parent._uriTable;
        _elementNameTable = parent._elementNameTable;
        _attributeNameTable = parent._attributeNameTable;
        _defaultNs = parent._defaultNs;
        _declSeen = false;
        DeclsOK = true;
      }

      /**
	 * Makes associated state become collectible,
	 * invalidating this context.
	 * {@link #setParent} must be called before
	 * this context may be used again.
	 */

      public void Clear() {
        _parent = null;
        _prefixTable = null;
        _uriTable = null;
        _elementNameTable = null;
        _attributeNameTable = null;
        _defaultNs = null;
      }

      /**
	 * Declare a Namespace prefix for this context.
	 *
	 * @param prefix The prefix to declare.
	 * @param uri The associated Namespace URI.
	 * @see org.xml.sax.helpers.NamespaceSupport#declarePrefix
	 */

      public void DeclarePrefix(string prefix, string uri) {
        // Lazy processing...
        if (!DeclsOK) {
          throw new InvalidOperationException("can't declare any more prefixes in this context");
        }
        if (!_declSeen) {
          CopyTables();
        }
        if (_declarations == null) {
          _declarations = new ArrayList();
        }

        prefix = string.Intern(prefix);
        uri = string.Intern(uri);
        if ("".Equals(prefix)) {
          if ("".Equals(uri)) {
            _defaultNs = null;
          } else {
            _defaultNs = uri;
          }
        } else {
          _prefixTable.Add(prefix, uri);
          _uriTable.Add(uri, prefix); // may wipe out another prefix
        }
        _declarations.Add(prefix);
      }

      /**
	 * Process an XML qualified name in this context.
	 *
	 * @param qName The XML qualified name.
	 * @param isAttribute true if this is an attribute name.
	 * @return An array of three strings containing the
	 *         URI part (or empty string), the local part,
	 *         and the raw name, all internalized, or null
	 *         if there is an undeclared prefix.
	 * @see org.xml.sax.helpers.NamespaceSupport#processName
	 */

      internal string[] ProcessName(string qName, bool isAttribute, bool namespaceDeclUris) {
        string[] name;
        Hashtable table;

        // detect errors in call sequence
        DeclsOK = false;

        // Select the appropriate table.
        if (isAttribute) {
          table = _attributeNameTable;
        } else {
          table = _elementNameTable;
        }

        // Start by looking in the cache, and
        // return immediately if the name
        // is already known in this content
        if (table.ContainsKey(qName)) {
          return (string[])table[qName];
        }

        // We haven't seen this name in this
        // context before.  Maybe in the parent
        // context, but we can't assume prefix
        // bindings are the same.
        name = new string[3];
        name[2] = string.Intern(qName);
        int index = qName.IndexOf(':');

        // No prefix.
        if (index == -1) {
          if (isAttribute) {
            if (qName == "xmlns" && namespaceDeclUris) {
              name[0] = NSDECL;
            } else {
              name[0] = "";
            }
          } else if (_defaultNs == null) {
            name[0] = "";
          } else {
            name[0] = _defaultNs;
          }
          name[1] = name[2];
        }
	    
          // Prefix
        else {
          string prefix = qName.Substring(0, index);
          string local = qName.Substring(index + 1);
          string uri = null;
          if ("".Equals(prefix)) {
            uri = _defaultNs;
          } else if (_prefixTable.ContainsKey(prefix)) {
            uri = (string)_prefixTable[prefix];
          }
          if (uri == null || (!isAttribute && "xmlns".Equals(prefix))) {
            return null;
          }
          name[0] = uri;
          name[1] = string.Intern(local);
        }

        // Save in the cache for future use.
        // (Could be shared with parent context...)
        table.Add(name[2], name);
        return name;
      }

      /**
	 * Look up the URI associated with a prefix in this context.
	 *
	 * @param prefix The prefix to look up.
	 * @return The associated Namespace URI, or null if none is
	 *         declared.	
	 * @see org.xml.sax.helpers.NamespaceSupport#getURI
	 */

      internal string GetURI(string prefix) {
        if ("".Equals(prefix)) {
          return _defaultNs;
        }
        if (_prefixTable == null) {
          return null;
        }
        if (_prefixTable.ContainsKey(prefix)) {
          return (string)_prefixTable[prefix];
        }
        return null;
      }

      /**
	 * Look up one of the prefixes associated with a URI in this context.
	 *
	 * <p>Since many prefixes may be mapped to the same URI,
	 * the return value may be unreliable.</p>
	 *
	 * @param uri The URI to look up.
	 * @return The associated prefix, or null if none is declared.
	 * @see org.xml.sax.helpers.NamespaceSupport#getPrefix
	 */

      internal string GetPrefix(string uri) {
        if (_uriTable == null) {
          return null;
        }
        if (_uriTable.ContainsKey(uri)) {
          return (string)_uriTable[uri];
        }
        return null;
      }

      /**
	 * Return an enumeration of prefixes declared in this context.
	 *
	 * @return An enumeration of prefixes (possibly empty).
	 * @see org.xml.sax.helpers.NamespaceSupport#getDeclaredPrefixes
	 */

      internal IEnumerable GetDeclaredPrefixes() {
        if (_declarations == null) {
          return Enumerable.Empty<object>();
        }
        return _declarations;
      }

      /**
	 * Return an enumeration of all prefixes currently in force.
	 *
	 * <p>The default prefix, if in force, is <em>not</em>
	 * returned, and will have to be checked for separately.</p>
	 *
	 * @return An enumeration of prefixes (never empty).
	 * @see org.xml.sax.helpers.NamespaceSupport#getPrefixes
	 */

      internal IEnumerable GetPrefixes() {
        if (_prefixTable == null) {
          return Enumerable.Empty<object>();
        }
        return _prefixTable.Keys;
      }

      ////////////////////////////////////////////////////////////////
      // Internal methods.
      ////////////////////////////////////////////////////////////////

      /**
	 * Copy on write for the internal tables in this context.
	 *
	 * <p>This class is optimized for the normal case where most
	 * elements do not contain Namespace declarations.</p>
	 */

      private void CopyTables() {
        if (_prefixTable != null) {
          _prefixTable = (Hashtable)_prefixTable.Clone();
        } else {
          _prefixTable = new Hashtable();
        }
        if (_uriTable != null) {
          _uriTable = (Hashtable)_uriTable.Clone();
        } else {
          _uriTable = new Hashtable();
        }
        _elementNameTable = new Hashtable();
        _attributeNameTable = new Hashtable();
        _declSeen = true;
      }

      ////////////////////////////////////////////////////////////////
      // Protected state.
      ////////////////////////////////////////////////////////////////
    }
  }

  // end of NamespaceSupport.java
}
