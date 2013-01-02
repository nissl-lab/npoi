using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class SingleXmlCellsDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_SingleXmlCells));
        CT_SingleXmlCells cells = null;

        public SingleXmlCellsDocument()
        {           
        }
        public SingleXmlCellsDocument(CT_SingleXmlCells cells)
        {
            this.cells = cells;
        }
        public static SingleXmlCellsDocument Parse(Stream stream)
        {
            CT_SingleXmlCells obj = (CT_SingleXmlCells)serializer.Deserialize(stream);
            return new SingleXmlCellsDocument(obj);
        }
        public CT_SingleXmlCells GetSingleXmlCells()
        {
            return cells;
        }
        public void SetSingleXmlCells(CT_SingleXmlCells cells)
        {
            this.cells = cells;
        }
        public void Save(Stream stream)
        {
            serializer.Serialize(stream, cells);
        }
    }
}
