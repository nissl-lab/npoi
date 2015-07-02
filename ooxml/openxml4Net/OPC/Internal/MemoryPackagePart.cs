using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.OpenXml4Net.OPC.Internal.Marshallers;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.Util;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    public class MemoryPackagePart : PackagePart
    {
        /**
         * Storage for the part data.
         */
        internal MemoryStream data;

        /**
         * Constructor.
         * 
         * @param pack
         *            The owner package.
         * @param partName
         *            The part name.
         * @param contentType
         *            The content type.
         * @throws InvalidFormatException
         *             If the specified URI is not OPC compliant.
         */
        public MemoryPackagePart(OPCPackage pack, PackagePartName partName,
                String contentType)
            : base(pack, partName, contentType)
        {

        }

        /**
         * Constructor.
         * 
         * @param pack
         *            The owner package.
         * @param partName
         *            The part name.
         * @param contentType
         *            The content type.
         * @param loadRelationships
         *            Specify if the relationships will be loaded.
         * @throws InvalidFormatException
         *             If the specified URI is not OPC compliant.
         */
        public MemoryPackagePart(OPCPackage pack, PackagePartName partName,
                String contentType, bool loadRelationships) :
            base(pack, partName, new ContentType(contentType), loadRelationships)
        {

        }

        protected override Stream GetInputStreamImpl()
        {
            // If this part has been created from scratch and/or the data buffer is
            // not
            // initialize, so we do it now.
            if (data == null)
            {
                return new MemoryStream();
            }
            MemoryStream newMs = new MemoryStream((int)data.Length);
            data.Position = 0;
            StreamHelper.CopyStream(data, newMs);
            newMs.Position = 0;
            return newMs;
        }

        protected override Stream GetOutputStreamImpl()
        {
            return new MemoryPackagePartOutputStream(this);
        }

        public override long Size
        {
            get
            {
                return data == null ? 0 : data.Length;
            }
        }
        public override void Clear()
        {
            data = null;
        }

        public override bool Save(Stream os)
        {
            return new ZipPartMarshaller().Marshall(this, os);
        }

        public override bool Load(Stream ios)
        {
            // Save it
            StreamHelper.CopyStream(ios, data);
            // All done
            return true;
        }

        public override void Close()
        {
            // Do nothing
        }

        public override void Flush()
        {
            // Do nothing
        }
    }
}
