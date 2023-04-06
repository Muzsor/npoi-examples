/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace CreateHyperlink
{
    class Program
    {
        static XWPFHyperlinkRun CreateHyperlinkRun(XWPFParagraph paragraph, String uri)
        {
            String rId = paragraph.Document.GetPackagePart().AddExternalRelationship(
              uri,
              XWPFRelation.HYPERLINK.Relation
             ).Id;

            return paragraph.CreateHyperlinkRun(rId);
        }
        static void Main(string[] args)
        {

            using (XWPFDocument doc = new XWPFDocument())
            {

                XWPFParagraph paragraph = doc.CreateParagraph();
                XWPFRun run = paragraph.CreateRun();
                run.SetText("This is a text paragraph having ");

                XWPFHyperlinkRun hyperlinkrun = CreateHyperlinkRun(paragraph, "https://www.google.com");
                hyperlinkrun.SetText("a link to Google");
                hyperlinkrun.SetColor("0000FF");
                hyperlinkrun.Underline =UnderlinePatterns.Single;

                run = paragraph.CreateRun();
                run.SetText(" in it.");
                using (FileStream out1 = new FileStream("hyperlink.docx", FileMode.Create))
                {
                    doc.Write(out1);
                }
            }
        }
    }
}
