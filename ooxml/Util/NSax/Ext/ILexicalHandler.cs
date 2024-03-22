namespace NSAX.Ext {
  /// <summary>
  ///   SAX2 extension handler for lexical events.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This is an optional extension handler for SAX2 to provide
  ///     lexical information about an XML document, such as comments
  ///     and CDATA section boundaries.
  ///     XML readers are not required to recognize this handler, and it
  ///     is not part of core-only SAX2 distributions.
  ///   </para>
  ///   <para>
  ///     The events in the lexical handler apply to the entire document,
  ///     not just to the document element, and all lexical handler events
  ///     must appear between the content handler's startDocument and
  ///     endDocument events.
  ///   </para>
  ///   <para>
  ///     To set the LexicalHandler for an XML reader, use the
  ///     <see cref="IXmlReader.SetProperty" /> method
  ///     with the property name
  ///     <c>http://xml.org/sax/properties/lexical-handler</c>
  ///     and an object implementing this interface (or null) as the value.
  ///     If the reader does not report lexical events, it will throw a
  ///     <see cref="SAXNotRecognizedException" />
  ///     when you attempt to register the handler.
  ///   </para>
  /// </summary>
  ////
  public interface ILexicalHandler {
    /// <summary>
    ///   Report the start of DTD declarations, if any.
    ///   <para>
    ///     This method is intended to report the beginning of the
    ///     DOCTYPE declaration; if the document has no DOCTYPE declaration,
    ///     this method will not be invoked.
    ///   </para>
    ///   <para>
    ///     All declarations reported through
    ///     <see cref="IDTDHandler" /> or
    ///     <see cref="IDeclHandler" /> events must appear
    ///     between the startDTD and <see cref="EndDTD" /> events.
    ///     Declarations are assumed to belong to the internal DTD subset
    ///     unless they appear between <see cref="StartEntity" />
    ///     and <see cref="EndEntity" /> events.  Comments and
    ///     processing instructions from the DTD should also be reported
    ///     between the startDTD and endDTD events, in their original
    ///     order of (logical) occurrence; they are not required to
    ///     appear in their correct locations relative to DTDHandler
    ///     or DeclHandler events, however.
    ///   </para>
    ///   <para>
    ///     Note that the start/endDTD events will appear within
    ///     the start/endDocument events from ContentHandler and
    ///     before the first
    ///     <see cref="IContentHandler.StartElement" />
    ///     event.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   The document type name.
    /// </param>
    /// <param name="publicId">
    ///   The declared public identifier for the
    ///   external DTD subset, or null if none was declared.
    /// </param>
    /// <param name="systemId">
    ///   The declared system identifier for the
    ///   external DTD subset, or null if none was declared.
    ///   (Note that this is not resolved against the document
    ///   base URI.)
    /// </param>
    /// <exception cref="SAXException">
    ///   The application may raise an
    ///   exception.
    /// </exception>
    /// <seealso cref="EndDTD" />
    /// <seealso cref="StartEntity" />
    void StartDTD(string name, string publicId, string systemId);

    /// <summary>
    ///   Report the end of DTD declarations.
    ///   <para>
    ///     This method is intended to report the end of the
    ///     DOCTYPE declaration; if the document has no DOCTYPE declaration,
    ///     this method will not be invoked.
    ///   </para>
    /// </summary>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    /// <seealso cref="StartDTD" />
    void EndDTD();

    /// <summary>
    ///   Report the beginning of some internal and external XML entities.
    ///   <para>
    ///     The reporting of parameter entities (including
    ///     the external DTD subset) is optional, and SAX2 drivers that
    ///     report LexicalHandler events may not implement it; you can use the
    ///     <c>http://xml.org/sax/features/lexical-handler/parameter-entities</c>
    ///     feature to query or control the reporting of parameter entities.
    ///   </para>
    ///   <para>
    ///     General entities are reported with their regular names,
    ///     parameter entities have '%' prepended to their names, and
    ///     the external DTD subset has the pseudo-entity name "[dtd]".
    ///   </para>
    ///   <para>
    ///     When a SAX2 driver is providing these events, all other
    ///     events must be properly nested within start/end entity
    ///     events.  There is no additional requirement that events from
    ///     <see cref="IDeclHandler" /> or
    ///     <see cref="IDTDHandler" /> be properly ordered.
    ///   </para>
    ///   <para>
    ///     Note that skipped entities will be reported through the
    ///     <see cref="IContentHandler.SkippedEntity" />
    ///     event, which is part of the ContentHandler interface.
    ///   </para>
    ///   <para>
    ///     Because of the streaming event model that SAX uses, some
    ///     entity boundaries cannot be reported under any
    ///     circumstances:
    ///   </para>
    ///   <ul>
    ///     <li>general entities within attribute values</li>
    ///     <li>parameter entities within declarations</li>
    ///   </ul>
    ///   <para>
    ///     These will be silently expanded, with no indication of where
    ///     the original entity boundaries were.
    ///   </para>
    ///   <para>
    ///     Note also that the boundaries of character references (which
    ///     are not really entities anyway) are not reported.
    ///   </para>
    ///   <para>All start/endEntity events must be properly nested.</para>
    /// </summary>
    /// <param name="name">
    ///   The name of the entity.  If it is a parameter
    ///   entity, the name will begin with '%', and if it is the
    ///   external DTD subset, it will be "[dtd]".
    /// </param>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    /// <seealso cref="EndEntity" />
    /// <seealso cref="IDeclHandler.InternalEntityDecl" />
    /// <seealso cref="IDeclHandler.ExternalEntityDecl" />
    void StartEntity(string name);

    /// <summary>
    ///   Report the end of an entity.
    /// </summary>
    /// <param name="name">
    ///   The name of the entity that is ending.
    /// </param>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    /// <seealso cref="StartEntity" />
    void EndEntity(string name);

    /// <summary>
    ///   Report the start of a CDATA section.
    ///   <para>
    ///     The contents of the CDATA section will be reported through
    ///     the regular <see cref="IContentHandler.Characters" /> event;
    ///     this event is intended only to report the boundary.
    ///   </para>
    /// </summary>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    /// <seealso cref="EndCDATA" />
    void StartCDATA();

    /// <summary>
    ///   Report the end of a CDATA section.
    /// </summary>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    /// <seealso cref="StartCDATA" />
    void EndCDATA();

    ////
    /// <summary>
    ///   Report an XML comment anywhere in the document.
    ///   <para>
    ///     This callback will be used for comments inside or outside the
    ///     document element, including comments in the external DTD
    ///     subset (if read).  Comments in the DTD must be properly
    ///     nested inside start/endDTD and start/endEntity events (if
    ///     used).
    ///   </para>
    /// </summary>
    /// <param name="ch">
    ///   An array holding the characters in the comment.
    /// </param>
    /// <param name="start">
    ///   The starting position in the array.
    /// </param>
    /// <param name="length">
    ///   The number of characters to use from the array.
    /// </param>
    /// <exception cref="SAXException">
    ///   The application may raise an exception.
    /// </exception>
    void Comment(char[] ch, int start, int length);
  }

  // end of LexicalHandler.java
}
