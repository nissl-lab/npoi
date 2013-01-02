using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class CalcChainDocument
    {
        CT_CalcChain calcChain;
        static XmlSerializer serializerObj = new XmlSerializer(typeof(CT_CalcChain));

        public CalcChainDocument()
        {
            this.calcChain = new CT_CalcChain();
        }
        internal CalcChainDocument(CT_CalcChain calcChain)
        {
            this.calcChain = calcChain;
        }

        public CT_CalcChain GetCalcChain()
        {
            return calcChain;
        }

        public void SetCalcChain(CT_CalcChain calcchain)
        {
            this.calcChain = calcchain;
        }

        public static CalcChainDocument Parse(Stream stream)
        {
            CT_CalcChain obj = (CT_CalcChain)serializerObj.Deserialize(stream);
            return new CalcChainDocument(obj);
        }

        public void Save(Stream stream)
        {
            serializerObj.Serialize(stream, calcChain);
        }

    }
}
