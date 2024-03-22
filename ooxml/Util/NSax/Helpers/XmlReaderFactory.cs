namespace NSAX.Helpers {
  using System;
  using System.Configuration;

  /// <summary>
  ///     Factory for creating an XML reader.
  ///     <blockquote>
  ///         <em>
  ///             This module, both source code and documentation, is in the
  ///             Public Domain, and comes with <strong>NO WARRANTY</strong>.
  ///         </em>
  ///         See <a href='http://www.saxproject.org'>http://www.saxproject.org</a>
  ///         for further information.
  ///     </blockquote>
  ///     <para>
  ///         This class contains static methods for creating an XML reader
  ///         from an explicit class name, or based on runtime defaults:
  ///     </para>
  ///     <code>
  ///     try {
  ///     IXmlReader myReader = XmlReaderFactory.CreateXmlReader();
  ///     } catch (SAXException ex) {
  ///       Console.WriteLine(ex.ToString());
  ///     }
  ///   </code>
  /// </summary>
  public sealed class XmlReaderFactory : BaseXmlReaderFactory {
    private static Lazy<IXmlReaderFactory> _current;
    private readonly SaxConfigurationSection _section;

    static XmlReaderFactory() {
      SetCurrent(LoadXmlReaderFactory);
    }

    private XmlReaderFactory(SaxConfigurationSection section) {
      _section = section;
    }

    /// <summary>
    ///     Gets the Current <see cref="IXmlReaderFactory" />
    /// </summary>
    public static IXmlReaderFactory Current {
      get { return _current.Value; }
    }

    /// <summary>
    ///     Sets the Current <see cref="IXmlReaderFactory" />
    /// </summary>
    /// <param name="factory">The factory</param>
    public static void SetCurrent(Func<IXmlReaderFactory> factory) {
      _current = new Lazy<IXmlReaderFactory>(factory);
    }

    public override IXmlReader CreateXmlReader() {
      return CreateXmlReader(_section.XmlReaderType);
    }

    private static IXmlReaderFactory LoadXmlReaderFactory() {
      var section = ConfigurationManager.GetSection("sax") as SaxConfigurationSection;
      if (section == null || string.IsNullOrEmpty(section.XmlReaderFactoryType)) {
        return new XmlReaderFactory(section);
      }
      Type factoryType = Type.GetType(section.XmlReaderFactoryType);
      if (factoryType == null) {
        throw new SAXException("SAX2 factory type " + section.XmlReaderFactoryType + " could not be found.");
      }
      try {
        return (IXmlReaderFactory)Activator.CreateInstance(factoryType);
      } catch (InvalidCastException ex) {
        throw new SAXException("SAX2 factory " + factoryType.FullName + " does not implement IXmlReaderFactory", ex);
      } catch (Exception ex) {
        throw new SAXException(
          "SAX2 factory " + factoryType.FullName + " loaded but cannot be instantiated (no empty public constructor?)",
          ex);
      }
    }
  }
}
