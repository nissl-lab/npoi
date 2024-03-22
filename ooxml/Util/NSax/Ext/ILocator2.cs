namespace NSAX.Ext {
  using System.Text;

  /// <summary>
  ///   SAX2 extension to augment the entity information provided
  ///   though a <see cref="ILocator" />.
  ///   If an implementation supports this extension, the Locator
  ///   provided in <see cref="IContentHandler.SetDocumentLocator" /> will implement this
  ///   interface, and the
  ///   <em>http://xml.org/sax/features/use-locator2</em> feature
  ///   flag will have the value <em>true</em>.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///   </blockquote>
  ///   <para>
  ///     XMLReader implementations are not required to support this
  ///     information, and it is not part of core-only SAX2 distributions.
  ///   </para>
  /// </summary>
  public interface ILocator2 : ILocator {
    ////
    /// <summary>
    ///   Returns the version of XML used for the entity.  This will
    ///   normally be the identifier from the current entity's
    ///   <em>&lt;?xml&amp;nbsp;version='...'&amp;nbsp;...?&gt;</em> declaration,
    ///   or be defaulted by the parser.
    /// </summary>
    /// <returns>
    ///   Identifier for the XML version being used to interpret
    ///   the entity's text, or null if that information is not yet
    ///   available in the current parsing state.
    /// </returns>
    string XmlVersion { get; }

    ////
    /// <summary>
    ///   Returns the name of the character encoding for the entity.
    ///   If the encoding was declared externally (for example, in a MIME
    ///   Content-Type header), that will be the name returned.  Else if there
    ///   was an <em>&lt;?xml&amp;nbsp;...encoding='...'?&gt;</em> declaration at
    ///   the start of the document, that encoding name will be returned.
    ///   Otherwise the encoding will been inferred (normally to be UTF-8, or
    ///   some UTF-16 variant), and that inferred name will be returned.
    ///   <para>
    ///     When an <see cref="InputSource" /> is used
    ///     to provide an entity's character stream, this method returns the
    ///     encoding provided in that input stream.
    ///   </para>
    ///   <para>
    ///     Note that some recent W3C specifications require that text
    ///     in some encodings be normalized, using Unicode Normalization
    ///     Form C, before processing.  Such normalization must be performed
    ///     by applications, and would normally be triggered based on the
    ///     value returned by this method.
    ///   </para>
    ///   <para>
    ///     Encoding names may be those used by the underlying JVM,
    ///     and comparisons should be case-insensitive.
    ///   </para>
    /// </summary>
    /// <returns>
    ///   Name of the character encoding being used to interpret
    ///   * the entity's text, or null if this was not provided for a *
    ///   character stream passed through an InputSource or is otherwise
    ///   not yet available in the current parsing state.
    /// </returns>
    Encoding Encoding { get; }
  }
}
