using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

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
        public bool Marshall(PackagePart part, Stream out1)
        {
            return part.Save(out1);
        }
    }
}
