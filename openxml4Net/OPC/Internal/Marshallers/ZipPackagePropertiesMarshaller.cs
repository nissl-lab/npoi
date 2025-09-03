using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
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

            // Build XML document (synchronous, in-memory operation)
            if (!BuildXmlDocument(part))
                return false;
                
            // Write XML to stream synchronously
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

    public override async Task<bool> MarshallAsync(PackagePart part, Stream out1, CancellationToken cancellationToken = default)
    {
        if (out1 is not ZipOutputStream zos)
        {
            throw new ArgumentException("ZipOutputStream expected!");
        }

        // Saving the part in the zip file
        string name = ZipHelper.GetZipItemNameFromOPCName(part.PartName.URI.ToString());
        ZipEntry ctEntry = new ZipEntry(name);

        try
        {
            // Save in ZIP
            zos.PutNextEntry(ctEntry); // Add entry in ZIP

            // Build XML document (synchronous, in-memory operation)
            if (!BuildXmlDocument(part))
                return false;

            // Save XML to stream asynchronously
            await StreamHelper.SaveXmlInStreamAsync(xmlDoc, zos, cancellationToken).ConfigureAwait(false);

            zos.CloseEntry();
            return true;
        }
        catch (IOException e)
        {
            throw new OpenXml4NetException(e.Message, e);
        }
        catch
        {
            return false;
        }
    }
}

}
