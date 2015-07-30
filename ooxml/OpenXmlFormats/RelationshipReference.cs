using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats
{
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", IsNullable = true)]
    public class ST_RelationshipId
    {
        //string _id;
        public ST_RelationshipId()
        { 
            
        }
        public static string NamespaceURI
        { 
            get{
                return "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
           }
        }

    }
}
