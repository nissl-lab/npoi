namespace NSAX.Ext {
  using NSAX.Helpers;

  /// <summary>
  ///   This class extends the SAX2 base handler class to support the
  ///   SAX2 <see cref="ILexicalHandler" />, {@link DeclHandler}, and
  ///   <see cref="IEntityResolver2" /> extensions.  Except for overriding the
  ///   original SAX1 <see cref="DefaultHandler.ResolveEntity" />
  ///   method the added handler methods just return.  Subclassers may
  ///   override everything on a method-by-method basis.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///   </blockquote>
  ///   <para>
  ///     <em>Note:</em> this class might yet learn that the
  ///     <em>ContentHandler.setDocumentLocator()</em> call might be passed a
  ///     <see cref="ILocator2" /> object, and that the
  ///     <em>ContentHandler.startElement()</em> call might be passed a
  ///     <see cref="IAttributes2" /> object.
  ///   </para>
  /// </summary>
  public class DefaultHandler2 : DefaultHandler, ILexicalHandler, IDeclHandler, IEntityResolver2 {
    public virtual void AttributeDecl(string eName, string aName, string type, string mode, string value) {
    }

    public virtual void ElementDecl(string name, string model) {
    }

    public virtual void ExternalEntityDecl(string name, string publicId, string systemId) {
    }

    public virtual void InternalEntityDecl(string name, string value) {
    }

    public virtual InputSource GetExternalSubset(string name, string baseUri) {
      return null;
    }

    public virtual InputSource ResolveEntity(string name, string publicId, string baseUri, string systemId) {
      return null;
    }

    public override InputSource ResolveEntity(string publicId, string systemId) {
      return ResolveEntity(null, publicId, null, systemId);
    }

    public virtual void StartCDATA() {
    }

    public virtual void EndCDATA() {
    }

    public virtual void StartDTD(string name, string publicId, string systemId) {
    }

    public virtual void EndDTD() {
    }

    public virtual void StartEntity(string name) {
    }

    public virtual void EndEntity(string name) {
    }

    public virtual void Comment(char[] ch, int start, int length) {
    }

    // SAX2 ext-1.0 DeclHandler
  }
}
