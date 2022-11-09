
using System;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.XWPF.UserModel;
using TestCases.OpenXml4Net;

namespace TestCases.OPC
{
    using NUnit.Framework;

    [TestFixture]
    public class TestPackageRelationship
    {
        [Test]
        public void PackageRelationshipNullSourceGetHashCode()
        {
            FileInfo targetFile = OpenXml4NetTestDataSamples.GetOutputFile("TestOpenPackageTMP.docx");

            FileInfo inputFile = OpenXml4NetTestDataSamples.GetSampleFile("TestOpenPackageINPUT.docx");
            
            
            // Copy the input file in the output directory
            FileHelper.CopyFile(inputFile.FullName, targetFile.FullName);
            

            // Create a namespace
            OPCPackage pkg = OPCPackage.OpenOrCreate(targetFile.FullName);
            var collection = new PackageRelationshipCollection(pkg);
            Uri baseUri = new Uri("ooxml://npoi.org"); //For test only.
            var relationShip = collection.AddRelationship(new Uri(baseUri,"/webSettings.xml"), TargetMode.Internal,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings", "rId4");

            var hashCode = relationShip.GetHashCode();
            Assert.NotZero(hashCode);
            
            pkg.Close();
            File.Delete(targetFile.FullName);
        }
    }
}
