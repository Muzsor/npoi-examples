/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI Examples: https://github.com/nissl-lab/npoi-examples
 * ==============================================================*/

using NPOI.XWPF.UserModel;
using System.IO;

namespace CreateEmptyDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            using (XWPFDocument doc = new XWPFDocument())
            {
                doc.CreateParagraph();

                using (FileStream sw = File.Create("blank.docx"))
                {
                    doc.Write(sw);
                }
            }
        }
    }
}
