using NPOI.XWPF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.IO;
using NPOI.OpenXml4Net.OPC;


namespace TestCases
{
    [TestFixture]
    public class TestIssue1353
    {
        
        [Test]
        public void TestOutputHasNorotWithShapeAttribute()
        {
            var samples = POIDataSamples.GetDocumentInstance();
            string inputPath = samples.GetFileInfo("issue1353from.docx").FullName;
            string projectDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory));
            string outputPath = Path.Combine(projectDir, "issue1353to.docx");
            string imagePath = samples.GetFileInfo("issue1353.png").FullName;
            using FileStream fs = new(inputPath, FileMode.Open, FileAccess.Read);
            XWPFDocument doc = new(fs);
            XWPFTable table = doc.CreateTable(1, 1);
            int row = 0;
            int col = 0;
            using FileStream picStream = new(imagePath, FileMode.Open, FileAccess.Read);
            string picName = "image.png";
            int widthEmus = 144 *  9525;
            int heightEmus = 144 * 9525;
            XWPFParagraph para = table.GetRow(row).GetCell(col).GetParagraphArray(0);
            para.Alignment = ParagraphAlignment.LEFT;
            para.SpacingAfter = 0;
            para.SpacingAfterLines = 0;
            para.SpacingBefore = 0;
            para.SpacingBeforeLines = 0;
            para.SpacingBetween = 1;
            XWPFRun r = para.CreateRun();
            using MemoryStream ms = new();
            picStream.Position = 0;
            picStream.CopyTo(ms);
            ms.Position = 0;
            XWPFPicture pic = r.AddPicture(ms, (int) PictureType.PNG, picName, widthEmus, heightEmus);
            
            
            pic.GetCTPicture().spPr.xfrm.rot = -90 * 60000;
            using(FileStream outFs = new(outputPath, FileMode.Create, FileAccess.Write))
            {
                doc.Write(outFs);
            }
            // check - Reopen the document and verify the XML content using XWPFDocument
            using (FileStream checkFs = new(outputPath, FileMode.Open, FileAccess.Read))
            {
                XWPFDocument checkDoc = new(checkFs);
                OPCPackage package = checkDoc.Package;
                
                // Get the document.xml part by iterating through parts
                PackagePart documentPart = null;
                foreach (PackagePart part in package.GetParts())
                {
                    if (part.PartName.Name.Equals("/word/document.xml", StringComparison.OrdinalIgnoreCase))
                    {
                        documentPart = part;
                        break;
                    }
                }
                
                ClassicAssert.IsNotNull(documentPart, "document.xml not found");
                
                using (Stream stream = documentPart.GetInputStream())
                using (StreamReader reader = new StreamReader(stream))
                {
                    string xmlContent = reader.ReadToEnd();
                    ClassicAssert.IsFalse(xmlContent.Contains("rotWithShape"),
                        "document.xml should not contain 'rotWithShape'");
                }
            }
            try
            {
                File.Delete(outputPath);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}