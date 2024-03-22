namespace NSAX.AElfred {
  using NSAX;
  using NSAX.Helpers;

  public class XmlReaderFactory :BaseXmlReaderFactory {
    /// <summary>
    /// Attempt to create an XMLReader from system defaults.
    /// </summary>
    /// <returns>
    /// A new XMLReader.
    /// </returns>
    /// <exception cref="T:Sax.Net.SAXException">If no default XMLReader class
    ///                 can be identified and instantiated.
    ///             </exception><seealso cref="M:Sax.Net.Helpers.BaseXmlReaderFactory.CreateXmlReader(System.String)"/>
    public override IXmlReader CreateXmlReader() {
      SAXDriver reader = new SAXDriver();
      reader.SetFeature("http://xml.org/sax/features/namespaces", true);;
      return reader;
    }
  }
}
