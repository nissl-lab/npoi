using System;
using NPOI;
using NPOI.POIFS.Crypt;
using NUnit.Framework;

namespace TestCases.POIFS.Crypt
{
    [TestFixture]
    public class TestCipherAlgorithm
    {
        [Test]
        public void Test()
        {
            Assert.AreEqual(128, CipherAlgorithm.aes128.defaultKeySize);

            foreach (CipherAlgorithm alg in CipherAlgorithm.Values)
            {
                Assert.AreEqual(alg, CipherAlgorithm.ValueOf(alg.ToString()));
            }

            Assert.AreEqual(CipherAlgorithm.aes128, CipherAlgorithm.FromEcmaId(0x660E));
            Assert.AreEqual(CipherAlgorithm.aes192, CipherAlgorithm.FromXmlId("AES", 192));

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
