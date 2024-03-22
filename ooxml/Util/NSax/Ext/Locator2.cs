namespace NSAX.Ext {
  using System.Text;

  using NSAX.Helpers;

  /// <summary>
  ///   SAX2 extension helper for holding additional Entity information,
  ///   implementing the <see cref="ILocator2" /> interface.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///   </blockquote>
  ///   <para> This is not part of core-only SAX2 distributions.</para>
  /// </summary>
  public class Locator2 : Locator, ILocator2 {
    private Encoding _encoding;
    private string _version;

    ////
    /// <summary>
    ///   Construct a new, empty Locator2 object.
    ///   This will not normally be useful, since the main purpose
    ///   of this class is to make a snapshot of an existing Locator.
    /// </summary>
    public Locator2() {
    }

    ////
    /// <summary>
    ///   Copy an existing Locator or Locator2 object.
    ///   If the object implements Locator2, values of the
    ///   <em>encoding</em> and <em>version</em>strings are copied,
    ///   otherwise they set to <em>null</em>.
    /// </summary>
    /// <param name="locator">
    ///   The existing Locator object.
    /// </param>
    public Locator2(ILocator locator) : base(locator) {
      if (locator is ILocator2) {
        var l2 = (ILocator2)locator;

        _version = l2.XmlVersion;
        _encoding = l2.Encoding;
      }
    }

    public virtual string XmlVersion {
      get { return _version; }
      set { _version = value; }
    }

    public virtual Encoding Encoding {
      get { return _encoding; }
      set { _encoding = value; }
    }
  }
}
