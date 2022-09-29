using NPOI.XWPF.UserModel;
using System;
using System.IO;

namespace ChangeOrientation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();

            var run = doc.CreateParagraph().CreateRun();
            run.SetText("Hello World!");

            doc.ChangeOrientation(NPOI.OpenXmlFormats.Wordprocessing.ST_PageOrientation.landscape);
            using (FileStream fs = new FileStream("test.docx", FileMode.Create))
            {
                doc.Write(fs);
                doc.Close();
            }
        }
    }
}
