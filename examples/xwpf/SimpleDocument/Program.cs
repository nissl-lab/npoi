using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.XWPF.UserModel;
using System.IO;

namespace SimpleDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();

            XWPFParagraph p1 = doc.CreateParagraph();
            p1.SetAlignment(ParagraphAlignment.CENTER);
            p1.SetBorderBottom(Borders.DOUBLE);
            p1.SetBorderTop(Borders.DOUBLE);

            p1.SetBorderRight(Borders.DOUBLE);
            p1.SetBorderLeft(Borders.DOUBLE);
            p1.SetBorderBetween(Borders.SINGLE);

            p1.SetVerticalAlignment(TextAlignment.TOP);

            XWPFRun r1 = p1.CreateRun();
            r1.SetBold(true);
            r1.SetText("The quick brown fox");
            r1.SetBold(true);
            r1.SetFontFamily("Courier");
            r1.SetUnderline(UnderlinePatterns.DotDotDash);
            r1.SetTextPosition(100);

            XWPFParagraph p2 = doc.CreateParagraph();
            p2.SetAlignment(ParagraphAlignment.RIGHT);

            //BORDERS
            p2.SetBorderBottom(Borders.DOUBLE);
            p2.SetBorderTop(Borders.DOUBLE);
            p2.SetBorderRight(Borders.DOUBLE);
            p2.SetBorderLeft(Borders.DOUBLE);
            p2.SetBorderBetween(Borders.SINGLE);

            XWPFRun r2 = p2.CreateRun();
            r2.SetText("jumped over the lazy dog");
            r2.SetStrike(true);
            r2.SetFontSize(20);

            XWPFRun r3 = p2.CreateRun();
            r3.SetText("and went away");
            r3.SetStrike(true);
            r3.SetFontSize(20);
            r3.SetSubscript(VerticalAlign.SUPERSCRIPT);


            XWPFParagraph p3 = doc.CreateParagraph();
            p3.SetWordWrap(true);
            p3.SetPageBreak(true);

            //p3.SetAlignment(ParagraphAlignment.DISTRIBUTE);
            p3.SetAlignment(ParagraphAlignment.BOTH);
            p3.SetSpacingLineRule(LineSpacingRule.EXACT);

            p3.SetIndentationFirstLine(600);


            XWPFRun r4 = p3.CreateRun();
            r4.SetTextPosition(20);
            r4.SetText("To be, or not to be: that is the question: "
                    + "Whether 'tis nobler in the mind to suffer "
                    + "The slings and arrows of outrageous fortune, "
                    + "Or to take arms against a sea of troubles, "
                    + "And by opposing end them? To die: to sleep; ");
            r4.AddBreak(BreakType.PAGE);
            r4.SetText("No more; and by a sleep to say we end "
                    + "The heart-ache and the thousand natural shocks "
                    + "That flesh is heir to, 'tis a consummation "
                    + "Devoutly to be wish'd. To die, to sleep; "
                    + "To sleep: perchance to dream: ay, there's the rub; "
                    + ".......");
            r4.SetItalic(true);
            //This would imply that this break shall be treated as a simple line break, and break the line after that word:

            XWPFRun r5 = p3.CreateRun();
            r5.SetTextPosition(-10);
            r5.SetText("For in that sleep of death what dreams may come");
            r5.AddCarriageReturn();
            r5.SetText("When we have shuffled off this mortal coil,"
                    + "Must give us pause: there's the respect"
                    + "That makes calamity of so long life;");
            r5.AddBreak();
            r5.SetText("For who would bear the whips and scorns of time,"
                    + "The oppressor's wrong, the proud man's contumely,");

            r5.AddBreak(BreakClear.ALL);
            r5.SetText("The pangs of despised love, the law's delay,"
                    + "The insolence of office and the spurns" + ".......");

            FileStream out1 = new FileStream("simple.docx", FileMode.Create);
            doc.Write(out1);
            out1.Close();
        }
    }
}
