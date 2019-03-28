using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.OpenXml4Net.OPC.Internal.Unmarshallers;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /**
     * Object implemented this interface are considered as part unmarshaller. A part
     * unmarshaller is responsible to unmarshall a part in order to load it from a
     * package.
     *
     * @author Julien Chable
     * @version 0.1
     */
    public interface PartUnmarshaller
    {

        /**
         * Save the content of the package in the stream
         *
         * @param in
         *            The input stream from which the part will be unmarshall.
         * @return The part freshly unmarshall from the input stream.
         * @throws OpenXml4NetException
         *             Throws only if any other exceptions are thrown by inner
         *             methods.
         */
        PackagePart Unmarshall(UnmarshallContext context, Stream in1);
    }
}
