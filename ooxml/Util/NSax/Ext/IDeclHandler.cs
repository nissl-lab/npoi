namespace NSAX.Ext {
  /// <summary>
  ///   SAX2 extension handler for DTD declaration events.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This is an optional extension handler for SAX2 to provide more
  ///     complete information about DTD declarations in an XML document.
  ///     XML readers are not required to recognize this handler, and it
  ///     is not part of core-only SAX2 distributions.
  ///   </para>
  ///   <para>
  ///     Note that data-related DTD declarations (unparsed entities and
  ///     notations) are already reported through the <see cref="IDTDHandler" /> interface.
  ///   </para>
  ///   <para>
  ///     If you are using the declaration handler together with a lexical
  ///     handler, all of the events will occur between the
  ///     <see cref="ILexicalHandler.StartDTD" /> and the
  ///     <see cref="ILexicalHandler.EndDTD" /> events.
  ///   </para>
  ///   <para>
  ///     To set the DeclHandler for an XML reader, use the
  ///     <see cref="IXmlReader.SetProperty" /> method
  ///     with the property name
  ///     <c>http://xml.org/sax/properties/declaration-handler</c>
  ///     and an object implementing this interface (or null) as the value.
  ///     If the reader does not report declaration events, it will throw a
  ///     <see cref="SAXNotRecognizedException" />
  ///     when you attempt to register the handler.
  ///   </para>
  /// </summary>
  public interface IDeclHandler {
    ////
    /// <summary>
    ///   Report an element type declaration.
    ///   <para>
    ///     The content model will consist of the string "EMPTY", the
    ///     string "ANY", or a parenthesised group, optionally followed
    ///     by an occurrence indicator.  The model will be normalized so
    ///     that all parameter entities are fully resolved and all whitespace
    ///     is removed,and will include the enclosing parentheses.  Other
    ///     normalization (such as removing redundant parentheses or
    ///     simplifying occurrence indicators) is at the discretion of the
    ///     parser.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   The element type name.
    /// </param>
    /// <param name="model">
    ///   The content model as a normalized string.
    /// </param>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    void ElementDecl(string name, string model);

    ////
    /// <summary>
    ///   Report an attribute type declaration.
    ///   <para>
    ///     Only the effective (first) declaration for an attribute will
    ///     be reported.  The type will be one of the strings "CDATA",
    ///     "ID", "IDREF", "IDREFS", "NMTOKEN", "NMTOKENS", "ENTITY",
    ///     "ENTITIES", a parenthesized token group with
    ///     the separator "|" and all whitespace removed, or the word
    ///     "NOTATION" followed by a space followed by a parenthesized
    ///     token group with all whitespace removed.
    ///   </para>
    ///   <para>
    ///     The value will be the value as reported to applications,
    ///     appropriately normalized and with entity and character
    ///     references expanded.
    ///   </para>
    /// </summary>
    /// <param name="eName">
    ///   The name of the associated element.
    /// </param>
    /// <param name="aName">
    ///   The name of the attribute.
    /// </param>
    /// <param name="type">
    ///   A string representing the attribute type.
    /// </param>
    /// <param name="mode">
    ///   A string representing the attribute defaulting mode
    ///   ("#IMPLIED", "#REQUIRED", or "#FIXED") or null if
    ///   none of these applies.
    /// </param>
    /// <param name="value">
    ///   A string representing the attribute's default value,
    ///   or null if there is none.
    /// </param>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    void AttributeDecl(string eName, string aName, string type, string mode, string value);

    /// <summary>
    ///   Report an internal entity declaration.
    ///   <para>
    ///     Only the effective (first) declaration for each entity
    ///     will be reported.  All parameter entities in the value
    ///     will be expanded, but general entities will not.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   The name of the entity.  If it is a parameter
    ///   entity, the name will begin with '%'.
    /// </param>
    /// <param name="value">
    ///   The replacement text of the entity.
    /// </param>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    /// <seealso cref="ExternalEntityDecl" />
    /// <seealso cref="IDTDHandler.UnparsedEntityDecl" />
    void InternalEntityDecl(string name, string value);

    /// <summary>
    ///   Report a parsed external entity declaration.
    ///   <para>
    ///     Only the effective (first) declaration for each entity
    ///     will be reported.
    ///   </para>
    ///   <para>
    ///     If the system identifier is a URL, the parser must resolve it
    ///     fully before passing it to the application.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   The name of the entity.  If it is a parameter
    ///   entity, the name will begin with '%'.
    /// </param>
    /// <param name="publicId">
    ///   The entity's public identifier, or null if none
    ///   was given.
    /// </param>
    /// <param name="systemId">
    ///   The entity's system identifier.
    /// </param>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    /// <seealso cref="InternalEntityDecl" />
    /// <seealso cref="IDTDHandler.UnparsedEntityDecl" />
    void ExternalEntityDecl(string name, string publicId, string systemId);
  }

  // end of DeclHandler.java
}
