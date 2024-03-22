namespace NSAX {
  using NSAX.Ext;
  using NSAX.Helpers;

  /// <summary>
  ///   Interface for a list of XML attributes.
  ///   <para></para>
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <see cref='http:///www.saxproject.org'>http:///www.saxproject.org</see>
  ///     for further information.
  ///   </blockquote>
  ///   <para></para>
  ///   <para>
  ///     This interface allows access to a list of attributes in
  ///     three different ways:
  ///   </para>
  ///   <para></para>
  ///   <ol>
  ///     <li>by attribute index;</li>
  ///     <li>by Namespace-qualified name; or</li>
  ///     <li>by qualified (prefixed) name.</li>
  ///   </ol>
  ///   <para></para>
  ///   <para>
  ///     The list will not contain attributes that were declared
  ///     #IMPLIED but not specified in the start tag.  It will also not
  ///     contain attributes used as Namespace declarations (xmlns///) unless
  ///     the <c>http:///xml.org/sax/features/namespace-prefixes</c>
  ///     feature is set to <c>true</c> (it is <c>false</c> by
  ///     default).
  ///     Because SAX2 conforms to the original "Namespaces in XML"
  ///     recommendation, it normally does not
  ///     give namespace declaration attributes a namespace URI.
  ///   </para>
  ///   <para></para>
  ///   <para>
  ///     Some SAX2 parsers may support using an optional feature flag
  ///     (<c>http:///xml.org/sax/features/xmlns-uris</c>) to request
  ///     that those attributes be given URIs, conforming to a later
  ///     backwards-incompatible revision of that recommendation.  (The
  ///     attribute's "local name" will be the prefix, or "xmlns" when
  ///     defining a default element namespace.)  For portability, handler
  ///     code should always resolve that conflict, rather than requiring
  ///     parsers that can change the setting of that feature flag.
  ///   </para>
  ///   <para></para>
  ///   <para>
  ///     If the namespace-prefixes feature (see above) is
  ///     <c>false</c>, access by qualified name may not be available; if
  ///     the <c>http:///xml.org/sax/features/namespaces</c> feature is
  ///     <c>false</c>, access by Namespace-qualified names may not be
  ///     available.
  ///   </para>
  ///   <para></para>
  ///   <para>
  ///     The order of attributes in the list is unspecified, and will
  ///     vary from implementation to implementation.
  ///   </para>
  ///   <para></para>
  /// </summary>
  /// <seealso cref="Attributes" />
  /// <seealso cref="IDeclHandler.AttributeDecl" />
  public interface IAttributes {
    /// <summary>
    ///   Gets the number of attributes in the list.
    ///   <para></para>
    ///   <para>
    ///     Once you know the number of attributes, you can iterate
    ///     through the list.
    ///   </para>
    ///   <para></para>
    /// </summary>
    /// <returns>The number of attributes in the list.</returns>
    /// <seealso cref="GetUri(int)" />
    /// <seealso cref="GetLocalName(int)" />
    /// <seealso cref="GetQName(int)" />
    /// <seealso cref="GetType(int)" />
    /// <seealso cref="GetValue(int)" />
    int Length { get; }

    /// <summary>
    ///   Look up an attribute's Namespace URI by index.
    /// </summary>
    /// <param name="index">
    ///   index The attribute index (zero-based).
    /// </param>
    /// <returns>
    ///   The Namespace URI, or the empty string if none
    ///   is available, or null if the index is out of range.
    /// </returns>
    /// <seealso cref="Length" />
    string GetUri(int index);

    /// <summary>
    ///   Look up an attribute's local name by index.
    /// </summary>
    /// <param name="index">The attribute index (zero-based)</param>
    /// <returns>
    ///   The local name, or the empty string if Namespace processing is not being performed, or null if the index is
    ///   out of range.
    /// </returns>
    /// <seealso cref="Length" />
    string GetLocalName(int index);

    /// <summary>
    ///   Look up an attribute's XML qualified (prefixed) name by index.
    /// </summary>
    /// <param name="index">The attribute index (zero-based).</param>
    /// <returns>
    ///   The XML qualified name, or the empty string
    ///   if none is available, or null if the index
    ///   is out of range.
    /// </returns>
    /// <seealso cref="Length" />
    string GetQName(int index);

    /// <summary>
    ///   Look up an attribute's type by index.
    ///   <para>
    ///     The attribute type is one of the strings "CDATA", "ID",
    ///     "IDREF", "IDREFS", "NMTOKEN", "NMTOKENS", "ENTITY", "ENTITIES",
    ///     or "NOTATION" (always in upper case).
    ///   </para>
    ///   <para>
    ///     If the parser has not read a declaration for the attribute,
    ///     or if the parser does not report attribute types, then it must
    ///     return the value "CDATA" as stated in the XML 1.0 Recommendation
    ///     (clause 3.3.3, "Attribute-Value Normalization").
    ///   </para>
    ///   <para>
    ///     For an enumerated attribute that is not a notation, the
    ///     parser will report the type as "NMTOKEN".
    ///   </para>
    /// </summary>
    /// <param name="index">The attribute index (zero-based).</param>
    /// <returns>
    ///   The attribute's type as a string, or null if the
    ///   index is out of range.
    /// </returns>
    /// <seealso cref="Length" />
    string GetType(int index);

    /// <summary>
    ///   Look up an attribute's value by index.
    ///   <para>
    ///     If the attribute value is a list of tokens (IDREFS,
    ///     ENTITIES, or NMTOKENS), the tokens will be concatenated
    ///     into a single string with each token separated by a
    ///     single space.
    ///   </para>
    /// </summary>
    /// <param name="index">The attribute index (zero-based).</param>
    /// <returns>
    ///   @return The attribute's value as a string, or null if the
    ///   index is out of range.
    /// </returns>
    /// <seealso cref="Length" />
    string GetValue(int index);

    
    /// <summary>
    ///   Look up the index of an attribute by Namespace name.
    /// </summary>
    /// <param name="uri">
    ///   The Namespace URI, or the empty string if
    ///   the name has no Namespace URI.
    /// </param>
    /// <param name="localName">The attribute's local name.</param>
    /// <returns>
    ///   The index of the attribute, or -1 if it does not
    ///   appear in the list.
    /// </returns>
    int GetIndex(string uri, string localName);

    /// <summary>
    ///   Look up the index of an attribute by XML qualified (prefixed) name.
    /// </summary>
    /// <param name="qName">The qualified (prefixed) name.</param>
    /// <returns>
    ///   The index of the attribute, or -1 if it does not
    ///   appear in the list.
    /// </returns>
    int GetIndex(string qName);

    /// <summary>
    ///   Look up an attribute's type by Namespace name.
    ///   <para>
    ///     See <seealso cref="GetType(int)" /> for a description
    ///     of the possible types.
    ///   </para>
    /// </summary>
    /// <param name="uri">
    ///   The Namespace URI, or the empty string if the
    ///   name has no Namespace URI.
    /// </param>
    /// <param name="localName">The local name of the attribute.</param>
    /// <returns>
    ///   The attribute type as a string, or null if the
    ///   attribute is not in the list or if Namespace
    ///   processing is not being performed.
    /// </returns>
    string GetType(string uri, string localName);

    /// <summary>
    ///   Look up an attribute's type by XML qualified (prefixed) name.
    ///   <para>
    ///     See <seealso cref="GetType(int)" /> for a description
    ///     * of the possible types.
    ///   </para>
    /// </summary>
    /// <param name="qName">The XML qualified name.</param>
    /// <returns>
    ///   The attribute type as a string, or null if the
    ///   attribute is not in the list or if qualified names
    ///   are not available.
    /// </returns>
    string GetType(string qName);

    /// <summary>
    ///   Look up an attribute's value by Namespace name.
    ///   <para>
    ///     See <see cref="GetValue(int)" /> for a description
    ///     of the possible values.
    ///   </para>
    /// </summary>
    /// <param name="uri">
    ///   The Namespace URI, or the empty string if the
    ///   name has no Namespace URI.
    /// </param>
    /// <param name="localName">The local name of the attribute.</param>
    /// <returns>
    ///   The attribute value as a string, or null if the
    ///   attribute is not in the list.
    /// </returns>
    string GetValue(string uri, string localName);

    /// <summary>
    ///   Look up an attribute's value by XML qualified (prefixed) name.
    ///   <para>
    ///     See <see cref="GetValue(int)"/> for a description
    ///     of the possible values.
    ///   </para>
    /// </summary>
    /// <param name="qName">The XML qualified name.</param>
    /// <returns>
    ///   The attribute value as a string, or null if the
    ///   attribute is not in the list or if qualified names
    ///   are not available.
    /// </returns>
    string GetValue(string qName);
  }

  // end of Attributes.java
}
