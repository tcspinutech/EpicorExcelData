using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpicorExcelData.Models
{
    internal class Result
    {
        internal int TotalRowsProcessed { get; set; }
        internal int TotalRowsImported { get; set; }
        internal int TotalRowsNotImported { get; set; }
        /// <summary>
        /// Key is row number. String list contains validation errors for that row number.
        /// </summary>
        internal List<Validation> Validations { get; }

        internal Result(int totalRowsProcessed, int totalRowsImported, int totalRowsNotImported,
            List<Validation> validations)
        {
            TotalRowsProcessed = totalRowsProcessed;
            TotalRowsImported = totalRowsImported;
            TotalRowsNotImported = totalRowsNotImported;
            Validations = validations;
        }
    }
}
