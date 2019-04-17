using System;
using System.IO;
using WordReader;

namespace WordReaderTest
{
    class Program
    {
        static void Main(string[] args)
        {
            docParser doc = new docParser("test.doc");
            if (doc.docIsOK == true)
            {
                string str = doc.getText();
                File.WriteAllText("DOCtest.txt", str);
            }

            docxParser docx = new docxParser("test.docx");
            if (docx.docxIsOK == true)
            {
                string str = docx.getText();
                File.WriteAllText("DOCXtest.txt", str);
            }

            Console.WriteLine("Press any key...");
            Console.ReadKey(true);
        }
    }
}
