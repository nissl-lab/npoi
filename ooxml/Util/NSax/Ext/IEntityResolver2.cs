namespace NSAX.Ext {
  using System.IO;

  /// <summary>
  ///   Extended interface for mapping external entity references to input
  ///   sources, or providing a missing external subset.  The
  ///   <see cref="IXmlReader.EntityResolver" /> property
  ///   is used to provide implementations of this interface to parsers.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///   </blockquote>
  ///   <para>
  ///     If a SAX application requires the customized handling which this
  ///     interface defines for external entities, it must ensure that it uses
  ///     an XMLReader with the
  ///     <em>http://xml.org/sax/features/use-entity-resolver2</em> feature flag
  ///     set to <em>true</em> (which is its default value when the feature is
  ///     recognized).  If that flag is unrecognized, or its value is false,
  ///     or the resolver does not implement this interface, then only the
  ///     <see cref="IEntityResolver" /> method will be used.
  ///   </para>
  ///   <para>
  ///     That supports three categories of application that modify entity
  ///     resolution.  <em>Old Style</em> applications won't know about this interface;
  ///     they will provide an EntityResolver.
  ///     <em>Transitional Mode</em> provide an EntityResolver2 and automatically
  ///     get the benefit of its methods in any systems (parsers or other tools)
  ///     supporting it, due to polymorphism.
  ///     Both <em>Old Style</em> and <em>Transitional Mode</em> applications will
  ///     work with any SAX2 parser.
  ///     <em>New style</em> applications will fail to run except on SAX2 parsers
  ///     that support this particular feature.
  ///     They will insist that feature flag have a value of "true", and the
  ///     EntityResolver2 implementation they provide  might throw an exception
  ///     if the original SAX 1.0 style entity resolution method is invoked.
  ///   </para>
  /// </summary>
  /// <seealso cref="IXmlReader.EntityResolver" />
  public interface IEntityResolver2 : IEntityResolver {
    /// <summary>
    ///   Allows applications to provide an external subset for documents
    ///   that don't explicitly define one.  Documents with DOCTYPE declarations
    ///   that omit an external subset can thus augment the declarations
    ///   available for validation, entity processing, and attribute processing
    ///   (normalization, defaulting, and reporting types including ID).
    ///   This augmentation is reported
    ///   through the <see cref="ILexicalHandler.StartDTD" /> method as if
    ///   the document text had originally included the external subset;
    ///   this callback is made before any internal subset data or errors
    ///   are reported.
    ///   <para>
    ///     This method can also be used with documents that have no DOCTYPE
    ///     declaration.  When the root element is encountered,
    ///     but no DOCTYPE declaration has been seen, this method is
    ///     invoked.  If it returns a value for the external subset, that root
    ///     element is declared to be the root element, giving the effect of
    ///     splicing a DOCTYPE declaration at the end the prolog of a document
    ///     that could not otherwise be valid.  The sequence of parser callbacks
    ///     in that case logically resembles this:
    ///   </para>
    ///   <code>
    ///     ... comments and PIs from the prolog (as usual)
    ///     StartDTD ("rootName", source.PublicId, source.SystemId);
    ///     StartEntity ("[dtd]");
    ///     ... declarations, comments, and PIs from the external subset
    ///     EndEntity ("[dtd]");
    ///     EndDTD ();
    ///     ... then the rest of the document (as usual)
    ///     StartElement (..., "rootName", ...);
    ///   </code>
    ///   <para>
    ///     Note that the InputSource gets no further resolution.
    ///     Implementations of this method may wish to invoke
    ///     <see cref="ResolveEntity" /> to gain benefits such as use
    ///     of local caches of DTD entities.  Also, this method will never be
    ///     used by a (non-validating) processor that is not including external
    ///     parameter entities.
    ///   </para>
    ///   <para>
    ///     Uses for this method include facilitating data validation when
    ///     interoperating with XML processors that would always require
    ///     undesirable network accesses for external entities, or which for
    ///     other reasons adopt a "no DTDs" policy.
    ///     Non-validation motives include forcing documents to include DTDs so
    ///     that attributes are handled consistently.
    ///     For example, an XPath processor needs to know which attibutes have
    ///     type "ID" before it can process a widely used type of reference.
    ///   </para>
    ///   <para>
    ///     <strong>Warning:</strong> Returning an external subset modifies
    ///     the input document.  By providing definitions for general entities,
    ///     it can make a malformed document appear to be well formed.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   Identifies the document root element.  This name comes
    ///   from a DOCTYPE declaration (where available) or from the actual
    ///   root element.
    /// </param>
    /// <param name="baseUri">
    ///   The document's base URI, serving as an additional
    ///   hint for selecting the external subset.  This is always an absolute
    ///   URI, unless it is null because the XMLReader was given an InputSource
    ///   without one.
    /// </param>
    /// <returns>
    ///   An InputSource object describing the new external subset
    ///   to be used by the parser, or null to indicate that no external
    ///   subset is provided.
    /// </returns>
    /// <exception cref="SAXException">
    ///   Any SAX exception, possibly wrapping
    ///   another exception.
    /// </exception>
    /// <exception cref="IOException">
    ///   Probably indicating a failure to create
    ///   a new InputStream or Reader, or an illegal URL.
    /// </exception>
    InputSource GetExternalSubset(string name, string baseUri);

    ////
    /// <summary>
    ///   Allows applications to map references to external entities into input
    ///   sources, or tell the parser it should use conventional URI resolution.
    ///   This method is only called for external entities which have been
    ///   properly declared.
    ///   This method provides more flexibility than the <see cref="IEntityResolver" />
    ///   interface, supporting implementations of more complex catalogue
    ///   schemes such as the one defined by the
    ///   <a href="http://www.oasis-open.org/committees/entity/spec-2001-08-06.html">OASIS XML Catalogs</a> specification.
    ///   <para>
    ///     Parsers configured to use this resolver method will call it
    ///     to determine the input source to use for any external entity
    ///     being included because of a reference in the XML text.
    ///     That excludes the document entity, and any external entity returned
    ///     by <see cref="GetExternalSubset" />.
    ///     When a (non-validating) processor is configured not to include
    ///     a class of entities (parameter or general) through use of feature
    ///     flags, this method is not invoked for such entities.
    ///   </para>
    ///   <para>
    ///     Note that the entity naming scheme used here is the same one
    ///     used in the <see cref="ILexicalHandler" />, or in the {@link
    ///     org.xml.sax.ContentHandler#skippedEntity
    ///     ContentHandler.skippedEntity()}
    ///     method.
    ///   </para>
    /// </summary>
    /// <param name="name">
    ///   Identifies the external entity being resolved.
    ///   Either "[dtd]" for the external subset, or a name starting
    ///   with "%" to indicate a parameter entity, or else the name of
    ///   a general entity.  This is never null when invoked by a SAX2
    ///   parser.
    /// </param>
    /// <param name="publicId">
    ///   The public identifier of the external entity being
    ///   referenced (normalized as required by the XML specification), or
    ///   null if none was supplied.
    /// </param>
    /// <param name="baseUri">
    ///   The URI with respect to which relative systemIDs
    ///   are interpreted.  This is always an absolute URI, unless it is
    ///   null (likely because the XMLReader was given an InputSource without
    ///   one).  This URI is defined by the XML specification to be the one
    ///   associated with the "&lt;" starting the relevant declaration.
    /// </param>
    /// <param name="systemId">
    ///   The system identifier of the external entity
    ///   being referenced; either a relative or absolute URI.
    ///   This is never null when invoked by a SAX2 parser; only declared
    ///   entities, and any external subset, are resolved by such parsers.
    /// </param>
    /// <returns>
    ///   An InputSource object describing the new input source to
    ///   be used by the parser.  Returning null directs the parser to
    ///   resolve the system ID against the base URI and open a connection
    ///   to resulting URI.
    /// </returns>
    /// <exception cref="SAXException">
    ///   Any SAX exception, possibly wrapping
    ///   another exception.
    /// </exception>
    /// <exception cref="IOException">
    ///   Probably indicating a failure to create
    ///   a new InputStream or Reader, or an illegal URL.
    /// </exception>
    InputSource ResolveEntity(string name, string publicId, string baseUri, string systemId);
  }
}
