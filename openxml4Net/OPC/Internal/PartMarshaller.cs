using System;
using System.IO;
using System.Text;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /**
     * Object implemented this interface are considered as part marshaller. A part
     * marshaller is responsible to marshall a part in order to be save in a
     * package.
     *
     * @author Julien Chable
     * @version 0.1
     */
    public interface PartMarshaller {

	    /**
	     * Save the content of the package in the stream
	     *
	     * @param part
	     *            Part to marshall.
	     * @param out
	     *            The output stream into which the part will be marshall.
	     * @return false if any marshall error occurs, else <b>true</b>
	     * @throws OpenXml4NetException
	     *             Throws only if any other exceptions are thrown by inner
	     *             methods.
	     */
        bool Marshall(PackagePart part, Stream out1);
    }
}
