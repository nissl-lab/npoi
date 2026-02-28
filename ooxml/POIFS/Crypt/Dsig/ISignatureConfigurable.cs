using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.POIFS.Crypt.Dsig
{
    public interface ISignatureConfigurable
    {
        void SetSignatureConfig(SignatureConfig signatureConfig);
    }
}
