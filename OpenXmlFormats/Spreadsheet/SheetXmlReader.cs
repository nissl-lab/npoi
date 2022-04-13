using Org.System.Xml.Sax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class SheetXmlReader : IXMLReader
    {
        public IEntityResolver EntityResolver { get; set; }
        public IDTDHandler DTDHandler { get; set; }
        public IContentHandler ContentHandler { get; set; }
        public IErrorHandler ErrorHandler { get; set; }

        public bool GetFeature(string name)
        {
            throw new NotImplementedException();
        }

        public IProperty GetProperty(string name)
        {
            throw new NotImplementedException();
        }

        public void Parse(InputSource input)
        {
            throw new NotImplementedException();
        }

        public void Parse(string systemId)
        {
            throw new NotImplementedException();
        }

        public void SetFeature(string name, bool value)
        {
            throw new NotImplementedException();
        }
    }
}
