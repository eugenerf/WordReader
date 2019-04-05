using System;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReader
{
    class Program
    {
        static void Main(string[] args)
        {
            docParser dp = new docParser("test2.doc");
            string str = dp.getText();
            File.WriteAllText("test.txt", str, Encoding.Unicode);

            Console.WriteLine("Any key to exit...");
            Console.ReadKey(true);
        }
    }
}
