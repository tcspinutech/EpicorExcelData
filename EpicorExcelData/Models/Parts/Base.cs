using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpicorExcelData.Models.Parts
{
    internal class Base : IPart
    {
        public int FromRowNumber { get; set; }
        internal string Number { get; set; }
        internal int Length { get; set; }
        internal decimal WebPrice { get; set; }
    }
}
