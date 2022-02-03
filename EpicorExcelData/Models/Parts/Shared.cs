using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EpicorExcelData.Models.Parts
{
    internal class Shared : IPart
    {
        public int FromRowNumber { get; set; }
        internal string EpicorPartNumber { get; set; }
        internal string WebSitePartNumber { get; set; }
        internal string MarketingProductName { get; set; }
        internal string ThreadAngle { get; set; }
        internal string ThreadClass { get; set; }
        internal string InternalThreadClass { get; set; }
        internal string ThreadType { get; set; }
        internal decimal? ANominalDiameterInches { get; set; }
        internal decimal? ANominalDiameterMillimeters { get; set; }
        internal decimal? RRootDiameterMinimumInches { get; set; }
        internal decimal? RRootDiameterMinimumMillimeters { get; set; }
        internal string DiameterCode { get; set; }
        internal string ImperialOrMetric { get; set; }
        internal decimal? LeadInches { get; set; }
        internal decimal? LeadMillimeters { get; set; }
        internal string LeadCode { get; set; }
        internal decimal? PitchInches { get; set; }
        internal decimal? PitchMillimeters { get; set; }
        internal int? Starts { get; set; }
        internal decimal? TurnsPerInch { get; set; }
        internal decimal? ThreadsPerMillimeters { get; set; }
        internal string Type123 { get; set; }
        internal string EndCodeForType4 { get; set; }
        internal int? NutSize { get; set; }
        internal string HandRightOrLeft { get; set; }
        internal string ScrewMaterial { get; set; }
        internal string DiaUnit { get; set; }
        internal string LeadUnit { get; set; }
        internal string ColorCode { get; set; }
        internal string AcmeCode { get; set; }
        internal decimal? ScrewWeight { get; set; }
        internal string WebSalable { get; set; }
        internal decimal? LeadAccuracy { get; set; }
        internal string LeadTime { get; set; }
        internal string MarketingDescription { get; set; }
        internal string Category1 { get; set; }
        internal string Category2 { get; set; }
        internal string Category3 { get; set; }
        internal string Category4 { get; set; }
        internal string Category5 { get; set; }
        internal string CadLink { get; set; }
        internal string Document1 { get; set; }
        internal string Document2 { get; set; }
        internal string Document3 { get; set; }
        internal string Document4 { get; set; }
        internal string Document5 { get; set; }
        internal string VideoLink { get; set; }
        internal string Calculator { get; set; }
        internal string Image1 { get; set; }
        internal string Image2 { get; set; }
        internal string Image3 { get; set; }
        internal string Image4 { get; set; }
    }
}
