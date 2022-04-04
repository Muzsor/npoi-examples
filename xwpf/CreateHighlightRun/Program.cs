using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace CreateHighlightRun
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph paragraph = doc.CreateParagraph();
            XWPFRun run = paragraph.CreateRun();
            run.SetText("This is text with ");
            run = paragraph.CreateRun();
            run.SetText("background color");
            run.GetCTR().AddNewRPr().shd = new CT_Shd();
            CT_Shd cTShd = run.GetCTR().AddNewRPr().shd;
            cTShd.val = ST_Shd.clear;
            cTShd.color = "auto";
            cTShd.fill = "00FFFF";

            run = paragraph.CreateRun();
            run.SetText(" and this is ");
            run = paragraph.CreateRun();
            run.SetText("highlighted");
            run.GetCTR().AddNewRPr().highlight = new CT_Highlight();
            run.GetCTR().AddNewRPr().highlight.val = ST_HighlightColor.yellow;
            run = paragraph.CreateRun();
            run.SetText(" text.");

            using (FileStream fs = new FileStream("highlight.docx", FileMode.Create))
            {
                doc.Write(fs);
            }
        }
    }
}
