using System;
using System.Collections.Generic;
using System.Text;
using NPOI.OpenXml4Net.OPC;
using System.Net.Mime;

namespace NPOI.Examples.CreateBasicOOXMLFile
{
    class Program
    {
        static void Main(string[] args)
        {
            //create ooxml file in memory
            Package p = Package.Create();

            //create package parts
            PackagePartName pn1=new PackagePartName(new Uri("/a/abcd/e",UriKind.Relative),true);
            if (!p.ContainPart(pn1))
                p.CreatePart(pn1, MediaTypeNames.Text.Plain);

            PackagePartName pn2 = new PackagePartName(new Uri("/b/test.xml", UriKind.Relative), true);
            if (!p.ContainPart(pn2))
                p.CreatePart(pn2, MediaTypeNames.Text.Xml);

            //save file 
            p.Save("test.zip");

            //don't forget to close it
            p.Close();
        }
    }
}
