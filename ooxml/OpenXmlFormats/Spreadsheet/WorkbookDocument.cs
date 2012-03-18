using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class WorkbookDocument
    {
        static XmlSerializer serializer = new XmlSerializer(typeof(CT_Workbook));

        CT_Workbook workbook = null;
        public WorkbookDocument()
        { 
            
        }
        public static WorkbookDocument Parse(Stream stream)
        {
            CT_Workbook obj = (CT_Workbook)serializer.Deserialize(stream);

            return new WorkbookDocument(obj);
        }
        public WorkbookDocument(CT_Workbook workbook)
        {
            this.workbook = workbook;
        }
        public CT_Workbook Workbook
        {
            get 
            {
                return this.workbook;
            }
        }
        public void Save(Stream stream)
        {
            serializer.Serialize(stream, workbook);
        }
    }
}
