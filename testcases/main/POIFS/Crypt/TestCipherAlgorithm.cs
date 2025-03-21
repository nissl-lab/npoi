using System;
using NPOI;
using NPOI.POIFS.Crypt;
using NUnit.Framework;using NUnit.Framework.Legacy;

namespace TestCases.POIFS.Crypt
{
    [TestFixture]
    public class TestCipherAlgorithm
    {
        [Test]
        public void Test()
        {
            ClassicAssert.AreEqual(128, CipherAlgorithm.aes128.defaultKeySize);

            foreach (CipherAlgorithm alg in CipherAlgorithm.Values)
            {
                ClassicAssert.AreEqual(alg, CipherAlgorithm.ValueOf(alg.ToString()));
            }

            ClassicAssert.AreEqual(CipherAlgorithm.aes128, CipherAlgorithm.FromEcmaId(0x660E));
            ClassicAssert.AreEqual(CipherAlgorithm.aes192, CipherAlgorithm.FromXmlId("AES", 192));

            try
            {
                CipherAlgorithm.FromEcmaId(0);
                Assert.Fail("Should throw exception");
            }
            catch (EncryptedDocumentException)
            {
                // expected
            }

            try
            {
                CipherAlgorithm.FromXmlId("AES", 1);
                Assert.Fail("Should throw exception");
            }
            catch (EncryptedDocumentException)
            {
                // expected
            }

            try
            {
                CipherAlgorithm.FromXmlId("RC1", 0x40);
                Assert.Fail("Should throw exception");
            }
            catch (EncryptedDocumentException)
            {
                // expected
            }
        }
    }
}
