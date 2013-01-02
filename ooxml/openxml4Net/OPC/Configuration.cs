using System;
using System.Text;
using System.IO;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Storage class for configuration storage parameters.
     * TODO xml syntax checking is no longer done with DOM4j parser -> remove the schema or do it ?
     *
     * @author CDubettier, Julen Chable
     * @version 1.0
     */
    public class Configuration
    {
        // TODO configuration by default. should be clearly stated that it should be
        // changed to match installation path
        // as schemas dir is needed in runtime
        static private String pathForXmlSchema = Environment.CurrentDirectory
                + @"\" + "src" + @"\" + "schemas";

        public static String PathForXmlSchema
        {
            get
            {
                return pathForXmlSchema;
            }
            set 
            {
                Configuration.pathForXmlSchema = value; 
            }
        }
    }
}
