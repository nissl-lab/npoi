namespace NSAX.Helpers {
  using System;

  /// <summary>
  ///     XmlReaderFactory base class
  /// </summary>
  public abstract class BaseXmlReaderFactory : IXmlReaderFactory {
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
    public abstract IXmlReader CreateXmlReader();

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
    public IXmlReader CreateXmlReader(string typeName) {
      Type type = Type.GetType(typeName);
      if (type == null) {
        throw new SAXException("SAX2 driver type " + typeName + " could not be found.");
      }
      return CreateXmlReader(type);
    }

    /// <summary>
    ///     Creates a new <see cref="IXmlReader" /> based on the type passed in.
    /// </summary>
    /// <param name="type">The type</param>
    /// <returns>A new XML reader.</returns>
    /// <exception cref="SAXException">If the type cannot be loaded, instantiated, and cast to IXmlReader.</exception>
    protected virtual IXmlReader CreateXmlReader(Type type) {
      try {
        return (IXmlReader)Activator.CreateInstance(type);
      } catch (InvalidCastException ex) {
        throw new SAXException("SAX2 driver class " + type.FullName + " does not implement IXmlReader", ex);
      } catch (Exception ex) {
        throw new SAXException(
          "SAX2 driver class " + type.FullName + " loaded but cannot be instantiated (no empty public constructor?)", ex);
      }
    }
  }
}
