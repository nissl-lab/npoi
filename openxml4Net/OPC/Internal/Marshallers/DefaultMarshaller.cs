using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NPOI.OpenXml4Net.OPC.Internal.Marshallers
{
    /**
     * Default marshaller that specified that the part is responsible to marshall its content.
     *
     * @author Julien Chable
     * @version 1.0
     * @see PartMarshaller
     */
    public class DefaultMarshaller : PartMarshaller
    {

        /**
         * Save part in the output stream by using the save() method of the part.
         *
         * @throws OpenXml4NetException
         *             If any error occur.
         */
        public bool Marshall(PackagePart part, Stream stream)
        {
            return part.Save(stream);
        }

        public async Task<bool> MarshallAsync(PackagePart part, Stream stream, CancellationToken cancellationToken=default)
        {
            return await part.SaveAsync(stream, cancellationToken);
        }

    }
}
