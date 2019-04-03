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
            docParser dp = new docParser("test3.doc");

            Console.WriteLine("Any key to exit...");
            Console.ReadKey(true);
        }
    }
}
