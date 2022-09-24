using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.Model;
using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace CreateHeaderAndFooter
{
    class Program
    {
        static void Main(string[] args)
        {
            using (XWPFDocument document = new XWPFDocument())
            {

                // create header-footer
                XWPFHeaderFooterPolicy headerFooterPolicy = document.GetHeaderFooterPolicy();
                if (headerFooterPolicy == null) headerFooterPolicy = document.CreateHeaderFooterPolicy();

                // create header start
                XWPFHeader header = headerFooterPolicy.CreateHeader(XWPFHeaderFooterPolicy.DEFAULT);

                XWPFParagraph paragraph = header.CreateParagraph();
                paragraph.Alignment = (ParagraphAlignment.CENTER);

                XWPFRun run = paragraph.CreateRun();
                run.SetText("Header");

                // create footer start
                XWPFFooter footer = headerFooterPolicy.CreateFooter(XWPFHeaderFooterPolicy.DEFAULT);

                paragraph = footer.CreateParagraph();
                paragraph.Alignment = (ParagraphAlignment.CENTER);

                run = paragraph.CreateRun();
                run.SetText("Footer");

                CT_SectPr sectPr = document.Document.body.sectPr;
                CT_PageMar pageMar = sectPr.AddPageMar();
                pageMar.left = 720; //720 TWentieths of an Inch Point (Twips) = 720/20 = 36 pt = 36/72 = 0.5"
                pageMar.right = 720;
                pageMar.top = 1440; //1440 Twips = 1440/20 = 72 pt = 72/72 = 1"
                pageMar.bottom = 1440;
                pageMar.header = 908; //45.4 pt * 20 = 908 = 45.4 pt header from top
                pageMar.footer = 568; //28.4 pt * 20 = 568 = 28.4 pt footer from bottom

                using (var fs = File.Create("CreateWordHeaderFooterTopBottom.docx"))
                {
                    document.Write(fs);
                }
            }
            
        }
    }
}
