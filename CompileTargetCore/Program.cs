using PSWikiTable;
using System;
using System.Linq;

namespace CompileTargetCore
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
            System.Collections.Generic.List<string> results = cmdlet.Invoke().OfType<string>().ToList();
            foreach (string item in results)
            {
                Console.WriteLine(item);
            }
            Console.ReadLine();
        }
    }
}
