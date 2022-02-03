using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpicorExcelData.Models
{
    internal class Validation
    {
        internal string File { get; set; }
        internal int RowNumber { get; set; }
        internal string Cell { get; set; }
        internal List<string> Messages { get; set; }

        internal Validation(string file, int rowNumber, string cell, List<string> messages)
        {
            File = file;
            RowNumber = rowNumber;
            Cell = cell;
            Messages = messages;
        }
    }
}
