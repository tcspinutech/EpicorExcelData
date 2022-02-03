using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EpicorExcelData.Columns;
using EpicorExcelData.Models;
using EpicorExcelData.Models.Parts;
using static System.Decimal;
using static System.Int32;
using Alloy = EpicorExcelData.Models.Parts.Alloy;
using Lead = EpicorExcelData.Models.Parts.Lead;
using alloyColumns = EpicorExcelData.Columns.Alloy;
using leadColumns = EpicorExcelData.Columns.Lead;

namespace EpicorExcelData
{
    internal class Excel
    {
        private const string Alloy = "Alloy";
        private const string Lead = "Lead";

        private static readonly string AppPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

        internal (Result resultModel, string exportedFile) ProcessFiles()
        {
            string source = Path.Combine(AppPath, $@"Resources\{Alloy}File.xlsx");
            var stream = File.Open(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var doc = SpreadsheetDocument.Open(stream, false);
            var (resultModel, baseData, sharedData, alloyData, leadData) = PrepareListData(doc, Alloy);
            string exportedFile = ExportToExcel(baseData, sharedData, alloyData, leadData, Alloy);
            return (resultModel, string.Empty);

            //string source = Path.Combine(AppPath, $@"Resources\{Lead}File.xlsx");
            //var stream = File.Open(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //using var doc = SpreadsheetDocument.Open(stream, false);
            //var (resultModel, baseData, sharedData, alloyData, leadData) = PrepareListData(doc, Lead);
            //string exportedFile = ExportToExcel(baseData, sharedData, alloyData, leadData, Lead);
            //return (resultModel, string.Empty);
        }

        private static (Result, List<Base>, List<Shared>, List<Alloy>, List<Lead>) PrepareListData(SpreadsheetDocument doc, string fileName)
        {
            var baseData = new List<Base>();
            var sharedData = new List<Shared>();
            var alloyData = new List<Alloy>();
            var leadData = new List<Lead>();

            bool isAlloy = fileName.Equals(Alloy, StringComparison.InvariantCultureIgnoreCase);

            var validations = new List<Validation>();

            Debug.Assert(doc.WorkbookPart != null, "doc.WorkbookPart != null");
            Debug.Assert(doc.WorkbookPart.Workbook != null, "doc.WorkbookPart.Workbook != null");
            var worksheet =
                    doc.WorkbookPart.Workbook.GetFirstChild<Sheets>()
                        ?.Elements<Sheet>().FirstOrDefault();

            if (worksheet == null)
            {
                var validationMessages = new List<string> { "First sheet wasn't found in this file." };
                validations.Add(new Validation(fileName, -1, string.Empty, validationMessages));

                return (new Result(0, 0, 0, validations), baseData, sharedData, alloyData, leadData);
            }

            var workbookPart = doc.WorkbookPart;
            Debug.Assert(worksheet.Id != null, "worksheet.Id != null");
            var worksheetData = ((WorksheetPart)workbookPart
                .GetPartById(worksheet.Id))
                .Worksheet.GetFirstChild<SheetData>();

            Debug.Assert(worksheetData != null, nameof(worksheetData) + " != null");
            int rowCount = worksheetData.Count();
            // First row contains headers, so subtract for correct number of rows processed.
            int rowsProcessed = rowCount - 1;
            // First row contains headers, so start at row number 2.
            for (var rowNumber = 4; rowNumber <= rowCount; rowNumber++)
            {
                var validationMessages = new List<string>();

                var arrayColumns = Enum.GetValues(isAlloy ? 
                    typeof(alloyColumns.Parts) : typeof(leadColumns.Parts));

                foreach (int i in arrayColumns)
                {

                    string partColumn = Enum.GetName(isAlloy ? 
                        typeof(alloyColumns.Parts) : typeof(leadColumns.Parts), i);

                    string partNumber = CellValue(workbookPart, worksheet, partColumn, rowNumber);

                    if (string.IsNullOrEmpty(partNumber))
                    {
                        validationMessages.Add($"Part number for size {i} is missing.");
                    }

                    bool exists = validations.FirstOrDefault(x =>
                        x.File.Equals(fileName, StringComparison.InvariantCultureIgnoreCase)
                        && x.RowNumber == rowNumber
                        && x.Cell.Equals(partColumn, StringComparison.InvariantCultureIgnoreCase)) == null;

                    if (validationMessages.Any() && !exists)
                    {
                        validations.Add(new Validation(fileName, rowNumber, partColumn, validationMessages));
                    }

                    if (baseData.FirstOrDefault(x =>
                            x.Number.Equals(partNumber, StringComparison.InvariantCultureIgnoreCase)) != null) continue;

                    string priceColumn = Enum.GetName(isAlloy ? 
                        typeof(alloyColumns.Prices) : typeof(leadColumns.Prices), i);

                    string price = CellValue(workbookPart, worksheet, priceColumn, rowNumber);
                    if (!TryParse(price, out decimal convertedPrice))
                    {
                        validationMessages.Add($"WebPrice for size {i} is missing or invalid.");
                    }

                    exists = validations.FirstOrDefault(x =>
                        x.File.Equals(fileName, StringComparison.InvariantCultureIgnoreCase)
                        && x.RowNumber == rowNumber
                        && x.Cell.Equals(partNumber, StringComparison.InvariantCultureIgnoreCase)) == null;

                    if (validationMessages.Any() && !exists)
                    {
                        validations.Add(new Validation(fileName, rowNumber, priceColumn, validationMessages));
                        continue;
                    }

                    if (alloyData == null) throw new ArgumentNullException(nameof(alloyData));

                    var baseItem = new Base
                    {
                        FromRowNumber = rowNumber,
                        Number = partNumber,
                        Length = i,
                        WebPrice = Round(convertedPrice, 2, MidpointRounding.AwayFromZero)
                    };

                    baseData.Add(baseItem);
                }

                if (validationMessages.Any())
                {
                    continue;
                }

                string epicorPartNumber = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.EpicorPartNumber);
                string webSitePartNumber = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.WebSitePartNumber);
                string marketingProductName = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.MarketingProductName);
                string threadAngle = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ThreadAngle);
                string threadClass = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ThreadClass);
                string internalThreadClass = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.InternalThreadClass);
                string threadType = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ThreadType);
                decimal? aNominalDiameterInches = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ANominalDiameterInches));
                decimal? aNominalDiameterMillimeters = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ANominalDiameterMillimeters));
                decimal? rRootDiameterMinimumInches = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.RRootDiameterMinimumInches));
                decimal? rRootDiameterMinimumMillimeters = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.RRootDiameterMinimumMillimeters));
                string diameterCode = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.DiameterCode);
                string imperialOrMetric = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ImperialOrMetric);
                decimal? leadInches = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.LeadInches));
                decimal? leadMillimeters = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.LeadMillimeters));
                string leadCode = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.LeadCode);
                decimal? pitchInches = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.PitchInches));
                decimal? pitchMillimeters = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.PitchMillimeters));
                int? starts = ToInt(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Starts));
                decimal? turnsPerInch = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.TurnsPerInch));
                decimal? threadsPerMillimeters = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ThreadsPerMillimeters));
                string type123 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Type123);
                string endCodeForType4 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.EndCodeForType4);
                int? nutSize = ToInt(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.NutSize));
                string handRightOrLeft = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.HandRightOrLeft);
                string screwMaterial = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ScrewMaterial);
                string diaUnit = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.DiaUnit);
                string leadUnit = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.LeadUnit);
                string colorCode = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ColorCode);
                string acmeCode = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.AcmeCode);
                decimal? screwWeight = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.ScrewWeight));
                string webSalable = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.WebSalable);
                decimal? leadAccuracy = ToDecimal(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.LeadAccuracy));
                string leadTime = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.LeadTime);
                string marketingDescription = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.MarketingDescription);
                string category1 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Category1);
                string category2 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Category2);
                string category3 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Category3);
                string category4 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Category4);
                string category5 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Category5);
                string cadLink = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.CadLink);
                string document1 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Document1);
                string document2 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Document2);
                string document3 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Document3);
                string document4 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Document4);
                string document5 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Document5);
                string videoLink = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.VideoLink);
                string calculator = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Calculator);
                string image1 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Image1);
                string image2 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Image2);
                string image3 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Image3);
                string image4 = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, isAlloy, Description.Image4);

                var sharedItem = new Shared
                {
                    FromRowNumber = rowNumber,
                    EpicorPartNumber = epicorPartNumber,
                    WebSitePartNumber = webSitePartNumber,
                    MarketingProductName = marketingProductName,
                    ThreadAngle = threadAngle,
                    ThreadClass = threadClass,
                    InternalThreadClass = internalThreadClass,
                    ThreadType = threadType,
                    ANominalDiameterInches = aNominalDiameterInches,
                    ANominalDiameterMillimeters = aNominalDiameterMillimeters,
                    RRootDiameterMinimumInches = rRootDiameterMinimumInches,
                    RRootDiameterMinimumMillimeters = rRootDiameterMinimumMillimeters,
                    DiameterCode = diameterCode,
                    ImperialOrMetric = imperialOrMetric,
                    LeadInches = leadInches,
                    LeadMillimeters = leadMillimeters,
                    LeadCode = leadCode,
                    PitchInches = pitchInches,
                    PitchMillimeters = pitchMillimeters,
                    Starts = starts,
                    TurnsPerInch = turnsPerInch,
                    ThreadsPerMillimeters = threadsPerMillimeters,
                    Type123 = type123,
                    EndCodeForType4 = endCodeForType4,
                    NutSize = nutSize,
                    HandRightOrLeft = handRightOrLeft,
                    ScrewMaterial = screwMaterial,
                    DiaUnit = diaUnit,
                    LeadUnit = leadUnit,
                    ColorCode = colorCode,
                    AcmeCode = acmeCode,
                    ScrewWeight = screwWeight,
                    WebSalable = webSalable,
                    LeadAccuracy = leadAccuracy,
                    LeadTime = leadTime,
                    MarketingDescription = marketingDescription,
                    Category1 = category1,
                    Category2 = category2,
                    Category3 = category3,
                    Category4 = category4,
                    Category5 = category5,
                    CadLink = cadLink,
                    Document1 = document1,
                    Document2 = document2,
                    Document3 = document3,
                    Document4 = document4,
                    Document5 = document5,
                    VideoLink = videoLink,
                    Calculator = calculator,
                    Image1 = image1,
                    Image2 = image2,
                    Image3 = image3,
                    Image4 = image4,
                };

                sharedData.Add(sharedItem);

                switch (isAlloy)
                {
                    case true:
                    {
                        int? epicorGroup = ToInt(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorGroup));
                        string epicorGroupDescription = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorGroupDescription);
                        string epicorRevisionDescription = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorRevisionDescription);
                        int? epicorRevisions = ToInt(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorRevisions));
                        string epicorPartConfigurable = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorPartConfigurable);
                        string epicorNonStockedItem = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorNonStockedItem);
                        string epicorGroupSalesSite = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorGroupSalesSite);
                        string epicorClass = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorClass);
                        string epicorDescription = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.EpicorDescription);
                        int? costingLotSize = ToInt(CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.CostingLotSize));
                        string cgMarketingProductName = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, true, alloyColumns.CgMarketingProductName);

                        var outItem = new Alloy
                        {
                            FromRowNumber = rowNumber,
                            EpicorGroup = epicorGroup,
                            EpicorGroupDescription = epicorGroupDescription,
                            EpicorRevisionDescription = epicorRevisionDescription,
                            EpicorRevisions = epicorRevisions,
                            EpicorPartConfigurable = epicorPartConfigurable,
                            EpicorNonStockedItem = epicorNonStockedItem,
                            EpicorGroupSalesSite = epicorGroupSalesSite,
                            EpicorClass = epicorClass,
                            EpicorDescription = epicorDescription,
                            CostingLotSize = costingLotSize,
                            CgMarketingProductName = cgMarketingProductName
                        };

                        alloyData.Add(outItem);
                        break;
                    }
                    case false:
                    {
                        string threadTypeCheck = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, false,
                            leadColumns.ThreadTypeCheck);
                        string externalThreadClass = CellValueForEnumDescription(workbookPart, worksheet, rowNumber, false,
                            leadColumns.ExternalThreadClass);

                        var leadItem = new Lead
                        {
                            FromRowNumber = rowNumber,
                            ThreadTypeCheck = threadTypeCheck,
                            ExternalThreadClass = externalThreadClass
                        };

                        leadData.Add(leadItem);
                        break;
                    }
                }
            }

            return (new Result(rowsProcessed, alloyData.Count, validations.Count, validations), baseData, sharedData, alloyData, leadData);
        }

        private static string ExportToExcel(IEnumerable<Base> baseData, IReadOnlyCollection<Shared> sharedData, IReadOnlyCollection<Alloy> alloyData, IReadOnlyCollection<Lead> leadData, string fileType)
        {
            bool isAlloy = fileType.Equals(Alloy, StringComparison.InvariantCultureIgnoreCase);
            string sheetName = isAlloy ? Alloy : Lead;
            var fileName = $"{fileType}.xlsx";
            string destination = Path.Combine(AppPath, fileName);

            if (File.Exists(destination))
            {
                File.Delete(destination);
            }

            using var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            var workbookPart = workbook.AddWorkbookPart();
            workbook.WorkbookPart.Workbook = new Workbook
            {
                Sheets = new Sheets()
            };

            uint sheetId = 1;

            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            sheetPart.Worksheet = new Worksheet(sheetData);

            var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

            Debug.Assert(sheets != null, nameof(sheets) + " != null");
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId =
                    sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

            var headerRow = new Row();

            var columns = new List<string>
            {
                "Part Number",
                "Length",
                "Web Price",
                Description.EpicorPartNumber,
                Description.WebSitePartNumber

            };

            if (isAlloy)
            {
                var alloySpecificColumns = new List<string>
                {
                    alloyColumns.EpicorGroup,
                    alloyColumns.EpicorGroupDescription,
                    alloyColumns.EpicorRevisionDescription,
                    alloyColumns.EpicorRevisions,
                    alloyColumns.EpicorPartConfigurable,
                    alloyColumns.EpicorNonStockedItem,
                    alloyColumns.EpicorGroupSalesSite,
                    alloyColumns.EpicorClass,
                    alloyColumns.EpicorDescription,
                    alloyColumns.CostingLotSize,
                    alloyColumns.CgMarketingProductName,
                    
                };

                columns.AddRange(alloySpecificColumns);
            }
            else
            {
                var leadSpecificColumns = new List<string>
                {
                    leadColumns.ThreadTypeCheck,
                    leadColumns.ExternalThreadClass
                            
                };

                columns.AddRange(leadSpecificColumns);
            }

            var appendSharedColumns = new List<string>
            {
                Description.ThreadAngle,
                Description.ThreadClass,
                Description.InternalThreadClass,
                Description.ThreadType,
                Description.ANominalDiameterInches,
                Description.ANominalDiameterMillimeters,
                Description.RRootDiameterMinimumInches,
                Description.RRootDiameterMinimumMillimeters,
                Description.DiameterCode,
                Description.ImperialOrMetric,
                Description.LeadInches,
                Description.LeadMillimeters,
                Description.LeadCode,
                Description.PitchInches,
                Description.PitchMillimeters,
                Description.Starts,
                Description.TurnsPerInch,
                Description.ThreadsPerMillimeters,
                Description.Type123,
                Description.EndCodeForType4,
                Description.NutSize,
                Description.HandRightOrLeft,
                Description.ScrewMaterial,
                Description.DiaUnit,
                Description.LeadUnit,
                Description.ColorCode,
                Description.AcmeCode,
                Description.ScrewWeight,
                Description.WebSalable,
                Description.LeadAccuracy,
                Description.LeadTime,
                Description.MarketingProductName,
                Description.MarketingDescription,
                Description.Category1,
                Description.Category2,
                Description.Category3,
                Description.Category4,
                Description.Category5,
                Description.CadLink,
                Description.Document1,
                Description.Document2,
                Description.Document3,
                Description.Document4,
                Description.Document5,
                Description.VideoLink,
                Description.Calculator,
                Description.Image1,
                Description.Image2,
                Description.Image3,
                Description.Image4
            };
                
            columns.AddRange(appendSharedColumns);

            foreach (var cell in columns.Select(column => new Cell
                     {
                         DataType = CellValues.String,
                         CellValue = new CellValue(column)
                     }))
            {
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);

            foreach (var baseItem in baseData)
            {
                var sharedItem = sharedData.FirstOrDefault(x => x.FromRowNumber == baseItem.FromRowNumber);
                if (sharedItem == null)
                {
                    throw new NullReferenceException(
                        $"Shared data could not be found for row number {baseItem.FromRowNumber}.");
                }

                var newRow = new Row();
                newRow.AppendChild(CreateCell(baseItem.Number));
                newRow.AppendChild(CreateCell(baseItem.Length));
                newRow.AppendChild(CreateCell(baseItem.WebPrice));
                newRow.AppendChild(CreateCell(sharedItem.EpicorPartNumber));
                newRow.AppendChild(CreateCell(sharedItem.WebSitePartNumber));

                if (isAlloy)
                {
                    var alloyItem = alloyData.FirstOrDefault(x => x.FromRowNumber == baseItem.FromRowNumber);
                    if (alloyItem == null)
                    {
                        throw new NullReferenceException(
                            $"Alloy data could not be found for row number {baseItem.FromRowNumber}.");
                    }

                    newRow.AppendChild(CreateCell(alloyItem.EpicorGroup));
                    newRow.AppendChild(CreateCell(alloyItem.EpicorGroupDescription));
                    newRow.AppendChild(CreateCell(alloyItem.EpicorRevisionDescription));
                    newRow.AppendChild(CreateCell(alloyItem.EpicorRevisions));
                    newRow.AppendChild(CreateCell(alloyItem.EpicorPartConfigurable));
                    newRow.AppendChild(CreateCell(alloyItem.EpicorNonStockedItem));
                    newRow.AppendChild(CreateCell(alloyItem.EpicorGroupSalesSite));
                    newRow.AppendChild(CreateCell(alloyItem.EpicorClass));
                    newRow.AppendChild(CreateCell(alloyItem.EpicorDescription));
                    newRow.AppendChild(CreateCell(alloyItem.CostingLotSize));
                    newRow.AppendChild(CreateCell(alloyItem.CgMarketingProductName));
                }
                else
                {
                    var leadItem = leadData.FirstOrDefault(x => x.FromRowNumber == baseItem.FromRowNumber);
                    if (leadItem == null)
                    {
                        throw new NullReferenceException(
                            $"Alloy data could not be found for row number {baseItem.FromRowNumber}.");
                    }   

                    newRow.AppendChild(CreateCell(leadItem.ThreadTypeCheck));
                    newRow.AppendChild(CreateCell(leadItem.ExternalThreadClass));
                }
                
                newRow.AppendChild(CreateCell(sharedItem.ThreadAngle));
                newRow.AppendChild(CreateCell(sharedItem.ThreadClass));
                newRow.AppendChild(CreateCell(sharedItem.InternalThreadClass));
                newRow.AppendChild(CreateCell(sharedItem.ThreadType));
                newRow.AppendChild(CreateCell(sharedItem.ANominalDiameterInches));
                newRow.AppendChild(CreateCell(sharedItem.ANominalDiameterMillimeters));
                newRow.AppendChild(CreateCell(sharedItem.RRootDiameterMinimumInches));
                newRow.AppendChild(CreateCell(sharedItem.RRootDiameterMinimumMillimeters));
                newRow.AppendChild(CreateCell(sharedItem.DiameterCode));
                newRow.AppendChild(CreateCell(sharedItem.ImperialOrMetric));
                newRow.AppendChild(CreateCell(sharedItem.LeadInches));
                newRow.AppendChild(CreateCell(sharedItem.LeadMillimeters));
                newRow.AppendChild(CreateCell(sharedItem.LeadCode));
                newRow.AppendChild(CreateCell(sharedItem.PitchInches));
                newRow.AppendChild(CreateCell(sharedItem.PitchMillimeters));
                newRow.AppendChild(CreateCell(sharedItem.Starts));
                newRow.AppendChild(CreateCell(sharedItem.TurnsPerInch));
                newRow.AppendChild(CreateCell(sharedItem.ThreadsPerMillimeters));
                newRow.AppendChild(CreateCell(sharedItem.Type123));
                newRow.AppendChild(CreateCell(sharedItem.EndCodeForType4));
                newRow.AppendChild(CreateCell(sharedItem.NutSize));
                newRow.AppendChild(CreateCell(sharedItem.HandRightOrLeft));
                newRow.AppendChild(CreateCell(sharedItem.ScrewMaterial));
                newRow.AppendChild(CreateCell(sharedItem.DiaUnit));
                newRow.AppendChild(CreateCell(sharedItem.LeadUnit));
                newRow.AppendChild(CreateCell(sharedItem.ColorCode));
                newRow.AppendChild(CreateCell(sharedItem.AcmeCode));
                newRow.AppendChild(CreateCell(sharedItem.ScrewWeight));
                newRow.AppendChild(CreateCell(sharedItem.WebSalable));
                newRow.AppendChild(CreateCell(sharedItem.LeadAccuracy));
                newRow.AppendChild(CreateCell(sharedItem.LeadTime));
                newRow.AppendChild(CreateCell(sharedItem.MarketingProductName));
                newRow.AppendChild(CreateCell(sharedItem.MarketingDescription));
                newRow.AppendChild(CreateCell(sharedItem.Category1));
                newRow.AppendChild(CreateCell(sharedItem.Category2));
                newRow.AppendChild(CreateCell(sharedItem.Category3));
                newRow.AppendChild(CreateCell(sharedItem.Category4));
                newRow.AppendChild(CreateCell(sharedItem.Category5));
                newRow.AppendChild(CreateCell(sharedItem.CadLink));
                newRow.AppendChild(CreateCell(sharedItem.Document1));
                newRow.AppendChild(CreateCell(sharedItem.Document2));
                newRow.AppendChild(CreateCell(sharedItem.Document3));
                newRow.AppendChild(CreateCell(sharedItem.Document4));
                newRow.AppendChild(CreateCell(sharedItem.Document5));
                newRow.AppendChild(CreateCell(sharedItem.VideoLink));
                newRow.AppendChild(CreateCell(sharedItem.Calculator));
                newRow.AppendChild(CreateCell(sharedItem.Image1));
                newRow.AppendChild(CreateCell(sharedItem.Image2));
                newRow.AppendChild(CreateCell(sharedItem.Image3));
                newRow.AppendChild(CreateCell(sharedItem.Image4));

                sheetData.AppendChild(newRow);
            }

            return destination;
        }

        // ReSharper disable AssignNullToNotNullAttribute
        // ReSharper disable PossibleNullReferenceException
        private static string CellValue(WorkbookPart workbookPart, Sheet worksheet, string spreadsheetColumn, int rowNumber)
        {
            var value = string.Empty;

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(worksheet.Id);
            var cellReference = $"{spreadsheetColumn}{rowNumber}";
            var cell = worksheetPart.Worksheet.Descendants<Cell>()
                .FirstOrDefault(x => x.CellReference == cellReference);

            //return string.IsNullOrEmpty(cell?.CellValue?.InnerText) ? string.Empty : cell.CellValue.InnerText;

            if (cell == null)
            {
                return value;
            }

            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                var stringId = Convert.ToInt32(cell.InnerText);
                value = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(stringId).InnerText;
            }
            else
            {
                if (cell.CellValue?.InnerText != null)
                {
                    value = cell.CellValue.InnerText;
                }
            }

            return value;
        }

        private static string CellValueForEnumDescription(WorkbookPart workbookPart, Sheet worksheet, int rowNumber, bool isAlloy, string description)
        {
            string columnLetter = isAlloy ? Letter.Fetch<alloyColumns.Append>(description) : Letter.Fetch<leadColumns.Append>(description);
            return CellValue(workbookPart, worksheet, columnLetter, rowNumber);
        }

        private static Cell CreateCell(dynamic value)
        {
            if (value == null)
            {
                return new Cell();
            }

            dynamic valueType = value.GetType().ToString();
            var cell = valueType switch
            {
                "System.Decimal" => new Cell
                {
                    DataType = CellValues.Number, CellValue = new CellValue(value.ToString("0.00"))
                },
                "System.Int32" => new Cell
                {
                    DataType = CellValues.Number, CellValue = new CellValue(value.ToString())
                },
                "System.String" => new Cell {DataType = CellValues.String, CellValue = new CellValue(value.ToString())},
                _ => throw new Exception($"{valueType} is not handled.")
            };

            return cell;
        }

        private static decimal? ToDecimal(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }

            _ = TryParse(value, out decimal convertedValue);
            return convertedValue;
        }

        private static int? ToInt(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }

            TryParse(value, out int convertedValue);
            return convertedValue;
        }
    }
}
