using System;
using System.IO;
using System.Text;

namespace WordReader
{
    class Program
    {
        static void Main(string[] args)
        {
            docxParser docx = new docxParser("test4.docx");

            string str = docx.getText();


            File.WriteAllText("test.txt", str, Encoding.Unicode);

            Console.WriteLine("Any key to exit...");
            Console.ReadKey(true);
        }        
    }
}
