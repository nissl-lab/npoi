using System;
using System.Collections.Generic;
using System.Text;
using NPOI.POIFS.FileSystem;
using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NPOI.HPSF;
using System.IO;

namespace CreateWordDoc
{
    class Program
    {
        static void Main(string[] args)
        {
            // POI apparently can't create a document from scratch,
            // so we need an existing empty dummy document
            POIFSFileSystem fs = new POIFSFileSystem(File.OpenRead("empty.doc"));
            HWPFDocument doc = new HWPFDocument(fs);

            // centered paragraph with large font size
            Range range = doc.GetRange();
            CharacterRun run1 = range.InsertAfter("one");
            //par1.SetSpacingAfter(200);
            //par1.SetJustification((byte)1);
            // justification: 0=left, 1=center, 2=right, 3=left and right

            //CharacterRun run1 = par1.InsertAfter("one");
            run1.SetFontSize(2 * 18);
            // font size: twice the point size

            // paragraph with bold typeface
            Paragraph par2 = run1.InsertAfter(new ParagraphProperties(), 0);
            par2.SetSpacingAfter(200);
            CharacterRun run2 = par2.InsertAfter("two two two two two two two two two two two two two");
            run2.SetBold(true);

            // paragraph with italic typeface and a line indent in the first line
            Paragraph par3 = run2.InsertAfter(new ParagraphProperties(), 0);
            par3.SetFirstLineIndent(200);
            par3.SetSpacingAfter(200);
            CharacterRun run3 = par3.InsertAfter("three three three three three three three three three "
                + "three three three three three three three three three three three three three three "
                + "three three three three three three three three three three three three three three");
            run3.SetItalic(true);

            // add a custom document property (needs POI 3.5; POI 3.2 doesn't save custom properties)
            DocumentSummaryInformation dsi = doc.DocumentSummaryInformation;
            CustomProperties cp = dsi.CustomProperties;
            if (cp == null)
                cp = new CustomProperties();
            cp.Put("myProperty", "foo bar baz");

            doc.Write(File.OpenWrite("new-hwpf-file.doc"));
        }
    }
}
