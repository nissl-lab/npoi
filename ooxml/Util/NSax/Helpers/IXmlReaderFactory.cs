namespace NSAX.Helpers {
  /// <summary> Factory for creating an XML reader.</summary>
  public interface IXmlReaderFactory {
    /// <summary>
    ///     Attempt to create an XMLReader from system defaults.
    /// </summary>
    /// <returns>
    ///     A new XMLReader.
    /// </returns>
    /// <exception cref="SAXException">
    ///     If no default XMLReader class
    ///     can be identified and instantiated.
    /// </exception>
    /// <seealso cref="CreateXmlReader(string)" />
    IXmlReader CreateXmlReader();

    /// <summary>
    ///     Attempt to create an XML reader from a type name.
    ///     <para>
    ///         Given a type name, this method attempts to load
    ///         and instantiate the class as an XML reader.
    ///     </para>
    ///     <para>
    ///         Note that this method will not be usable in environments where
    ///         the caller is not permitted to load types dynamically.
    ///     </para>
    /// </summary>
    /// <returns>
    ///     A new XML reader.
    /// </returns>
    /// <exception cref="SAXException">
    ///     If the type cannot be loaded, instantiated, and cast to IXmlReader.
    /// </exception>
    /// <seealso cref="CreateXmlReader()" />
    IXmlReader CreateXmlReader(string typeName);
  }
}