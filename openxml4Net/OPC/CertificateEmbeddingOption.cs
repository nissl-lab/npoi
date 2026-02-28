using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Specifies the location where the X.509 certificate that is used in signing is stored.
     *
     * @author Julien Chable
     */
    public enum CertificateEmbeddingOption
    {
        /** The certificate is embedded in its own PackagePart. */
        IN_CERTIFICATE_PART,
        /** The certificate is embedded in the SignaturePart that is created for the signature being added. */
        IN_SIGNATURE_PART,
        /** The certificate in not embedded in the package. */
        NOT_EMBEDDED
    }

}
