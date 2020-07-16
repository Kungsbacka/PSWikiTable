using PSWikiTable;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CompileTargetFramework
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ConvertToWikiTableCmdlet cmdlet = new ConvertToWikiTableCmdlet()
            {
                Path = "",
                Worksheet = null,
                NoFormatting = false,
                WikiBaseUri = new Uri("")
            };
            List<string> results = cmdlet.Invoke().OfType<string>().ToList();
            foreach (string item in results)
            {
                Console.WriteLine(item);
            }
            Console.ReadLine();
        }
    }
}
