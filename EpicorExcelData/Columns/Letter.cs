using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpicorExcelData.Columns
{
    internal static class Letter
    {
        internal static string Fetch<T>(string description)
        {
            foreach (var field in typeof(T).GetFields())
            {
                if (Attribute.GetCustomAttribute(field,
                        typeof(DescriptionAttribute)) is not DescriptionAttribute attribute) continue;
                if (attribute.Description == description)
                    return field.Name;
            }

            throw new ArgumentException(@"Not found.", nameof(description));
        }
    }
}
