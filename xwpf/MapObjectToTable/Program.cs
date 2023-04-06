using NPOI.XWPF.UserModel;
using System.IO;

namespace MapObjectToTable
{
    class Program
    {
        static void Main()
        {
            using (XWPFDocument document = new XWPFDocument())
            {
                XWPFTableWrapper<ExampleClass> newWrapper = new XWPFTableWrapper<ExampleClass>(document);

                ExampleClass exampleData = new ExampleClass()
                {
                    Enum = ExampleEnum.Red,
                    Name = "This is a test",
                    IgnoredMember = "This member will be ignored",
                    Address = new ExampleChildClass()
                    {
                        Address = "Burgemeester Schönfeldplein",
                        City = "Winschoten",
                        CountryCode = "NL"
                    }
                };
                newWrapper.AddRow(exampleData);

                using FileStream fs = new FileStream("test.docx", FileMode.Create);
                document.Write(fs);
            }
        }
    }
}
