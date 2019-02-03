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
		if (!(out1 is ZipOutputStream)) {
			throw new ArgumentException("ZipOutputStream expected!");
		}
		ZipOutputStream zos = (ZipOutputStream) out1;

		// Saving the part in the zip file
		string name = ZipHelper
				.GetZipItemNameFromOPCName(part.PartName.URI.ToString());
        ZipEntry ctEntry = new ZipEntry(name);

        try
        {
            // Save in ZIP
            zos.PutNextEntry(ctEntry); // Add entry in ZIP

            base.Marshall(part, out1); // Marshall the properties inside a XML
            // Document
            StreamHelper.SaveXmlInStream(xmlDoc, out1);

            zos.CloseEntry();
        }
        catch (IOException e)
        {
            throw new OpenXml4NetException(e.Message);
        }
        catch
        {
            return false; 
        }
		return true;
	}
}

}
