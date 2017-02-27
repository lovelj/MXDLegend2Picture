using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MXDLegendExport
{
    static class StringExtension
    {
        public static string Convert2Str(this string[] args, string separator)
        {
             if (args == null || args.Length == 0)
                return string.Empty;
            StringBuilder result = new StringBuilder();
            int count = 0;
            foreach (var item in args)
            {
                count++;
                if (item == null)
                    continue;
                result.Append(item.ToString() + (count == args.Length ? null : separator));
            }
            return result.ToString();
        }
    }
}
