using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.OpenXml4Net.OPC.Internal.Marshallers;
using ICSharpCode.SharpZipLib.Zip;

namespace NPOI.OpenXml4Net.OPC
{
/**
 * Zip implementation of a PackagePart.
 * 
 * @author Julien Chable
 * @version 1.0
 * @see PackagePart
 */
public class ZipPackagePart : PackagePart {

    /**
     * The zip entry corresponding to this part.
     */
    private ZipEntry zipEntry;

    /**
     * Constructor.
     * 
     * @param container
     *            The container package.
     * @param partName
     *            Part name.
     * @param contentType
     *            Content type.
     * @throws InvalidFormatException
     *             Throws if the content of this part invalid.
     */
    public ZipPackagePart(OPCPackage container, PackagePartName partName,
            String contentType):base(container, partName, contentType)
    {
        
    }

    /**
     * Constructor.
     * 
     * @param container
     *            The container package.
     * @param zipEntry
     *            The zip entry corresponding to this part.
     * @param partName
     *            The part name.
     * @param contentType
     *            Content type.
     * @throws InvalidFormatException
     *             Throws if the content of this part is invalid.
     */
    public ZipPackagePart(OPCPackage container, ZipEntry zipEntry,
            PackagePartName partName, String contentType):	base(container, partName, contentType)
    {
    
        this.zipEntry = zipEntry;
    }

    /**
     * Get the zip entry of this part.
     * 
     * @return The zip entry in the zip structure coresponding to this part.
     */
    public ZipEntry ZipArchive
    {
        get{
        return zipEntry;
        }
    }

    /**
     * Implementation of the getInputStream() which return the inputStream of
     * this part zip entry.
     * 
     * @return Input stream of this part zip entry.
     */
    
    protected override Stream GetInputStreamImpl()
    {
        // We use the getInputStream() method from java.util.zip.ZipFile
        // class which return an InputStream to this part zip entry.
        return ((ZipPackage) _container).ZipArchive
                .GetInputStream(zipEntry);
    }

    protected override Stream GetOutputStreamImpl()
    {
        return null;
    }

    public override long Size
    {
        get
        {
            return zipEntry.Size;
        }
    }

    public override bool Save(Stream os){
        return new ZipPartMarshaller().Marshall(this, os);
    }


    public override bool Load(Stream ios)
    {
        throw new InvalidOperationException("Method not implemented !");
    }


    public override void Close()
    {
        throw new InvalidOperationException("Method not implemented !");
    }


    public override void Flush()
    {
        throw new InvalidOperationException("Method not implemented !");
    }
}
}
