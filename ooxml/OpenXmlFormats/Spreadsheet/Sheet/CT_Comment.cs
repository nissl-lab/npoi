using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "comment")]
    public class CT_Comment
    {

        private CT_Rst textField = new CT_Rst(); // required element 

        private string refField = string.Empty; // required attribute

        private uint authorIdField = 0; // required attribute

        private string guidField = null; // optional attribute

        //public CT_Comment()
        //{
        //    this.textField = new CT_Rst();
        //}

        [XmlElement("text")]
        public CT_Rst text
        {
            get
            {
                return this.textField;
            }
            set
            {
                this.textField = value;
            }
        }

        [XmlAttribute("ref")]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }

        [XmlAttribute("authorId")]
        public uint authorId
        {
            get
            {
                return this.authorIdField;
            }
            set
            {
                this.authorIdField = value;
            }
        }

        [XmlAttribute("guid")] // 0..1 TODO: Type is ST_Guid
        public string guid
        {
            get
            {
                return this.guidField;
            }
            set
            {
                this.guidField = value;
            }
        }
    }
}
