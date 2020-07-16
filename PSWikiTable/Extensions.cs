using System;

namespace PSWikiTable
{
    internal static class Extensions
    {
        public static bool IEquals(this string left, string right)
        {
            return left.Equals(right, StringComparison.OrdinalIgnoreCase);
        }
    }
}
