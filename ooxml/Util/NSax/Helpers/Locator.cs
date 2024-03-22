namespace NSAX.Helpers {
  /// <summary>
  ///   Provide an optional convenience implementation of Locator.
  ///   <blockquote>
  ///     <em>
  ///       This module, both source code and documentation, is in the
  ///       Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///     </em>
  ///     See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///     for further information.
  ///   </blockquote>
  ///   <para>
  ///     This class is available mainly for application writers, who
  ///     can use it to make a persistent snapshot of a locator at any
  ///     point during a document parse:
  ///   </para>
  ///   <code>
  ///     ILocator locator;
  ///     ILocator startloc;
  ///     public void setLocator (ILocator locator)
  ///     {
  ///       this.locator = locator;
  ///     }
  ///     public void StartDocument ()
  ///     {
  ///       ILocator startloc = new Locator(locator);
  ///     }
  ///   </code>
  ///   <para>
  ///     Normally, parser writers will not use this class, since it
  ///     is more efficient to provide location information only when
  ///     requested, rather than constantly updating a Locator object.
  ///   </para>
  /// </summary>
  /// <seealso cref="ILocator" />
  public class Locator : ILocator {
    private readonly int _columnNumber;
    private readonly int _lineNumber;
    private readonly string _publicId;
    private readonly string _systemId;

    ////
    /// <summary>
    ///   Zero-argument constructor.
    ///   <para>
    ///     This will not normally be useful, since the main purpose
    ///     of this class is to make a snapshot of an existing Locator.
    ///   </para>
    /// </summary>
    public Locator() {
    }

    /// <summary>
    ///   Copy constructor.
    ///   <para>
    ///     Create a persistent copy of the current state of a locator.
    ///     When the original locator changes, this copy will still keep
    ///     the original values (and it can be used outside the scope of
    ///     DocumentHandler methods).
    ///   </para>
    /// </summary>
    /// <param name="locator">The locator to copy.</param>
    public Locator(ILocator locator) {
      _publicId = locator.PublicId;
      _systemId = locator.SystemId;
      _lineNumber = locator.LineNumber;
      _columnNumber = locator.ColumnNumber;
    }

    public string PublicId {
      get { return _publicId; }
    }

    public string SystemId {
      get { return _systemId; }
    }

    public int LineNumber {
      get { return _lineNumber; }
    }

    public int ColumnNumber {
      get { return _columnNumber; }
    }
  }

  // end of LocatorImpl.java
}
