namespace NSAX.Helpers {
  using System.Configuration;

  public class SaxConfigurationSection : ConfigurationSection {
    [ConfigurationProperty("xmlReaderType", IsRequired = false)]
    public string XmlReaderType {
      get { return (string)this["xmlReaderType"]; }
    }

    [ConfigurationProperty("xmlReaderFactoryType")]
    public string XmlReaderFactoryType {
      get { return (string)this["xmlReaderFactoryType"]; }
    }
  }
}
