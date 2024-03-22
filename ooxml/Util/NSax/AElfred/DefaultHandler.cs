namespace NSAX.AElfred {
  using NSAX.Ext;

  /// <summary>
  ///   This class extends the SAX base handler class to support the
  ///   SAX2 Lexical and Declaration handlers.  All the handler methods do
  ///   is return; except that the SAX base class handles fatal errors by
  ///   throwing an exception.
  /// </summary>
  public class DefaultHandler : NSAX.Helpers.DefaultHandler, ILexicalHandler, IDeclHandler {
    /// <summary>
    ///   called on attribute declarations
    /// </summary>
    /// <param name="element"></param>
    /// <param name="name"></param>
    /// <param name="type"></param>
    /// <param name="defaultType"></param>
    /// <param name="defaltValue"></param>
    public virtual void AttributeDecl(string element, string name, string type, string defaultType, string defaltValue) {
    }

    /// <summary>
    ///   called on element declarations
    /// </summary>
    /// <param name="name"></param>
    /// <param name="model"></param>
    public virtual void ElementDecl(string name, string model) {
    }

    /// <summary>
    ///   called on external entity declarations
    /// </summary>
    /// <param name="name"></param>
    /// <param name="pubid"></param>
    /// <param name="sysid"></param>
    public virtual void ExternalEntityDecl(string name, string pubid, string sysid) {
    }

    /// <summary>
    ///   called on internal entity declarations
    /// </summary>
    /// <param name="name"></param>
    /// <param name="value"></param>
    public virtual void InternalEntityDecl(string name, string value) {
    }

    /// <summary>
    ///   called before parsing CDATA characters
    /// </summary>
    public virtual void StartCDATA() {
    }

    /// <summary>
    ///   called after parsing CDATA characters
    /// </summary>
    public virtual void EndCDATA() {
    }

    /// <summary>
    ///   called when the doctype is partially parsed
    /// </summary>
    /// <param name="root"></param>
    /// <param name="pubid"></param>
    /// <param name="sysid"></param>
    public virtual void StartDTD(string root, string pubid, string sysid) {
    }

    /// <summary>
    ///   called after the doctype is parsed
    /// </summary>
    public virtual void EndDTD() {
    }

    /// <summary>
    /// called before parsing a general entity in content
    /// </summary>
    /// <param name="name"></param>
    public virtual void StartEntity(string name) {
    }

    /// <summary>
    ///   called after parsing a general entity in content
    /// </summary>
    /// <param name="name"></param>
    public virtual void EndEntity(string name) {
    }

    /// <summary>
    ///   called when comments are parsed
    /// </summary>
    /// <param name="buf"></param>
    /// <param name="off"></param>
    /// <param name="len"></param>
    public virtual void Comment(char[] buf, int off, int len) {
    }
  }
}
