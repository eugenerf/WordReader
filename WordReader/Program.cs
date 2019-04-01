using System;
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
            docParser dp = new docParser("test1.doc");

            Console.WriteLine("Any key to exit...");
            Console.ReadKey(true);
        }

        static void ShowBytes(byte[] Src)
        {
            foreach (byte s in Src)
                Console.Write($"{s:X2}");
            Console.WriteLine();
        }
    }
}
