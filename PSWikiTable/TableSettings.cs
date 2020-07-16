using System;
using System.Collections.Generic;

namespace PSWikiTable
{
    internal class TableSettings
    {
        public int Width { get; set; }
        public int Height { get; set; }
        public bool NoFormatting { get; set; }
        public Uri WikiBaseUri { get; set; }
        public Dictionary<string, string> Templates { get; set; }
    }
}
