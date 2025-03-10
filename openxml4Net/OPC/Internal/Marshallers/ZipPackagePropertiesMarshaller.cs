using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.OPC.Internal.Marshallers
{
/**
 * Package core properties marshaller specialized for zipped package.
 *
 * @author Julien Chable
 */
public class ZipPackagePropertiesMarshaller:PackagePropertiesMarshaller 
{
	public override bool Marshall(PackagePart part, Stream out1)
	{
		if (out1 is not ZipOutputStream zos) {
			throw new ArgumentException("ZipOutputStream expected!");
		}

        // Saving the part in the zip file
		string name = ZipHelper
				.GetZipItemNameFromOPCName(part.PartName.URI.ToString());
        ZipEntry ctEntry = new ZipEntry(name);

        try
        {
            // Save in ZIP
            zos.PutNextEntry(ctEntry); // Add entry in ZIP

            base.Marshall(part, zos); // Marshall the properties inside a XML
            // Document
            StreamHelper.SaveXmlInStream(xmlDoc, zos);

            zos.CloseEntry();
        }
        catch (IOException e)
        {
            throw new OpenXml4NetException(e.Message, e);
        }
        catch
        {
            return false; 
        }
		return true;
	}
}

}
