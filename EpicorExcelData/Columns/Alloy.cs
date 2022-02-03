using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpicorExcelData.Columns
{
    // ReSharper disable InconsistentNaming
    internal class Alloy : Description
    {
        /// <summary>
        /// Columns have base part in 1" increments. Index represent size.
        /// </summary>
        internal enum Parts
        {
            AU = 1,
            AV = 2,
            AW = 3,
            AX = 4,
            AY = 5,
            AZ = 6,
            BA = 7,
            BB = 8,
            BC = 9,
            BD = 10,
            BE = 11,
            BF = 12,
            BG = 13,
            BH = 14,
            BI = 15,
            BJ = 16,
            BK = 17,
            BL = 18,
            BM = 19,
            BN = 20,
            BO = 21,
            BP = 22,
            BQ = 23,
            BR = 24,
            BS = 25,
            BT = 26,
            BU = 27,
            BV = 28,
            BW = 29,
            BX = 30,
            BY = 31,
            BZ = 32,
            CA = 33,
            CB = 34,
            CC = 35,
            CD = 36,
            CE = 37,
            CF = 38,
            CG = 39,
            CH = 40,
            CI = 41,
            CJ = 42,
            CK = 43,
            CL = 44,
            CM = 45,
            CN = 46,
            CO = 47,
            CP = 48,
            CQ = 49,
            CR = 50,
            CS = 51,
            CT = 52,
            CU = 53,
            CV = 54,
            CW = 55,
            CX = 56,
            CY = 57,
            CZ = 58,
            DA = 59,
            DB = 60,
            DC = 61,
            DD = 62,
            DE = 63,
            DF = 64,
            DG = 65,
            DH = 66,
            DI = 67,
            DJ = 68,
            DK = 69,
            DL = 70,
            DM = 71,
            DN = 72
        }
        
        /// <summary>
        /// Prices. Index corresponds with part size.
        /// </summary>
        internal enum Prices
        {
            DQ = 1,
            DR = 2,
            DS = 3,
            DT = 4,
            DU = 5,
            DV = 6,
            DW = 7,
            DX = 8,
            DY = 9,
            DZ = 10,
            EA = 11,
            EB = 12,
            EC = 13,
            ED = 14,
            EE = 15,
            EF = 16,
            EG = 17,
            EH = 18,
            EI = 19,
            EJ = 20,
            EK = 21,
            EL = 22,
            EM = 23,
            EN = 24,
            EO = 25,
            EP = 26,
            EQ = 27,
            ER = 28,
            ES = 29,
            ET = 30,
            EU = 31,
            EV = 32,
            EW = 33,
            EX = 34,
            EY = 35,
            EZ = 36,
            FA = 37,
            FB = 38,
            FC = 39,
            FD = 40,
            FE = 41,
            FF = 42,
            FG = 43,
            FH = 44,
            FI = 45,
            FJ = 46,
            FK = 47,
            FL = 48,
            FM = 49,
            FN = 50,
            FO = 51,
            FP = 52,
            FQ = 53,
            FR = 54,
            FS = 55,
            FT = 56,
            FU = 57,
            FV = 58,
            FW = 59,
            FX = 60,
            FY = 61,
            FZ = 62,
            GA = 63,
            GB = 64,
            GC = 65,
            GD = 66,
            GE = 67,
            GF = 68,
            GG = 69,
            GH = 70,
            GI = 71,
            GJ = 72
        }

        internal const string EpicorGroup = "Epicor Group";
        internal const string EpicorGroupDescription = "Epicor Group Description";
        internal const string EpicorRevisionDescription = "Epicor Revision Description";
        internal const string EpicorRevisions = "Epicor Revisions";
        internal const string EpicorPartConfigurable = "Epicor Part - Configurable";
        internal const string EpicorNonStockedItem = "Epicor - Non Stocked Item";
        internal const string EpicorGroupSalesSite = "Epicor Group Sales Site";
        internal const string EpicorClass = "Epicor Class";
        internal const string EpicorDescription = "Epicor Description";
        internal const string CostingLotSize = "Costing Lot Size (Quantity Bearing, Type:  Manufactured)";
        internal const string CgMarketingProductName = "CG-Marketing Product Name";

        internal enum Append
        {
            [Description(EpicorPartNumber)]
            A,
            [Description(WebSitePartNumber)]
            B,
            [Description(EpicorGroup)]
            C,
            [Description(EpicorGroupDescription)]
            D,
            [Description(EpicorRevisionDescription)]
            E,
            [Description(EpicorRevisions)]
            F,
            [Description(EpicorPartConfigurable)]
            G,
            [Description(EpicorNonStockedItem)]
            H,
            [Description(EpicorGroupSalesSite)]
            I,
            [Description(EpicorClass)]
            J,
            [Description(EpicorDescription)]
            K,
            [Description(CostingLotSize)]
            L,
            [Description(CgMarketingProductName)]
            M,
            [Description(MarketingProductName)]
            N,
            [Description(ThreadAngle)]
            O,
            [Description(ThreadClass)]
            Q,
            [Description(InternalThreadClass)]
            R,
            [Description(ThreadType)]
            S,
            [Description(ANominalDiameterInches)]
            T,
            [Description(ANominalDiameterMillimeters)]
            U,
            [Description(RRootDiameterMinimumInches)]
            V,
            [Description(RRootDiameterMinimumMillimeters)]
            W,
            [Description(DiameterCode)]
            X,
            [Description(ImperialOrMetric)]
            Y,
            [Description(LeadInches)]
            Z,
            [Description(LeadMillimeters)]
            AA,
            [Description(LeadCode)]
            AB,
            [Description(PitchInches)]
            AC,
            [Description(PitchMillimeters)]
            AD,
            [Description(Starts)]
            AE,
            [Description(TurnsPerInch)]
            AF,
            [Description(ThreadsPerMillimeters)]
            AG,
            [Description(Type123)]
            AH,
            [Description(EndCodeForType4)]
            AI,
            [Description(NutSize)]
            AJ,
            [Description(HandRightOrLeft)]
            AK,
            [Description(ScrewMaterial)]
            AL,
            [Description(DiaUnit)]
            AM,
            [Description(LeadUnit)]
            AN,
            [Description(ColorCode)]
            AO,
            [Description(AcmeCode)]
            AP,
            [Description(ScrewWeight)]
            AQ,
            [Description(WebSalable)]
            AR,
            [Description(LeadAccuracy)]
            AS,
            [Description(LeadTime)]
            AT,
            [Description(MarketingDescription)]
            GK,
            [Description(Category1)]
            GL,
            [Description(Category2)]
            GM,
            [Description(Category3)]
            GN,
            [Description(Category4)]
            GO,
            [Description(Category5)]
            GP,
            [Description(CadLink)]
            GQ,
            [Description(Document1)]
            GR,
            [Description(Document2)]
            GS,
            [Description(Document3)]
            GT,
            [Description(Document4)]
            GU,
            [Description(Document5)]
            GV,
            [Description(VideoLink)]
            GW,
            [Description(Calculator)]
            GX,
            [Description(Image1)]
            GY,
            [Description(Image2)]
            GZ,
            [Description(Image3)]
            HA,
            [Description(Image4)]
            HB
        }
    }
}
