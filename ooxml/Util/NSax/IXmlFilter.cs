namespace NSAX {
  using NSAX.Helpers;

  /// <summary>
  ///   Interface for an XML filter.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     An XML filter is like an XML reader, except that it obtains its
  ///     events from another XML reader rather than a primary source like
  ///     an XML document or database.  Filters can modify a stream of
  ///     events as they pass on to the final application.
  ///   </para>
  ///   <para>
  ///     The XMLFilterImpl helper class provides a convenient base
  ///     for creating SAX2 filters, by passing on all <see cref="IEntityResolver" />,
  ///     <see cref="IContentHandler" /> and <see cref="IErrorHandler" /> events automatically.
  ///   </para>
  /// </summary>
  /// <seealso cref="XmlFilter" />
  public interface IXmlFilter : IXmlReader {
    /// <summary>
    ///   Get or sets the parent reader.
    ///   <para>
    ///     This method allows the application to query the parent
    ///     reader (which may be another filter).  It is generally a
    ///     bad idea to perform any operations on the parent reader
    ///     directly: they should all pass through this filter.
    ///   </para>
    ///   <para>
    ///     This method allows the application to link the filter to
    ///     a parent reader (which may be another filter).  The argument
    ///     may not be null.
    ///   </para>
    /// </summary>
    /// <returns>The parent filter, or null if none has been set.</returns>
    IXmlReader Parent { get; set; }
  }

  // end of XMLFilter.java
}
