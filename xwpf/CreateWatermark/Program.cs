using NPOI.XWPF.Model;
using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace CreateWatermark
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph paragraph = doc.CreateParagraph();
            XWPFRun run = paragraph.CreateRun();
            run.SetText("The Body:");
            var hfPolicy = doc.CreateHeaderFooterPolicy();
            hfPolicy.CreateWatermark("My Watermark");


            using (FileStream fs = new FileStream("watermark.docx", FileMode.Create))
            {
                doc.Write(fs);
            }
        }
    }
}
