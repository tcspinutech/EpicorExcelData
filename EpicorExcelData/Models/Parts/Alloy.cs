using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpicorExcelData.Models.Parts
{
    internal class Alloy : IPart
    {
        public int FromRowNumber { get; set; }
        internal int? EpicorGroup { get; set; }
        internal string EpicorGroupDescription { get; set; }
        internal string EpicorRevisionDescription { get; set; }
        internal int? EpicorRevisions { get; set; }
        internal string EpicorPartConfigurable { get; set; }
        internal string EpicorNonStockedItem { get; set; }
        internal string EpicorGroupSalesSite { get; set; }
        internal string EpicorClass { get; set; }
        internal string EpicorDescription { get; set; }
        internal int? CostingLotSize { get; set; }
        internal string CgMarketingProductName { get; set; }
    }
}
