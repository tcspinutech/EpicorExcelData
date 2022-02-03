using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpicorExcelData.Models.Parts
{
    internal class Lead : IPart
    {
        public int FromRowNumber { get; set; }
        internal string ThreadTypeCheck { get; set; }
        internal string ExternalThreadClass { get; set; }
    }   
}
