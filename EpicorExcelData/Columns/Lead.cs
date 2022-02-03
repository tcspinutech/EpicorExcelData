using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpicorExcelData.Columns
{
    // ReSharper disable InconsistentNaming
    internal class Lead : Description
    {
        /// <summary>
        /// Columns have base part in 1" increments. Index represent size.
        /// </summary>
        internal enum Parts
        {
            AJ = 1,
            AK = 2,
            AL = 3,
            AM = 4,
            AN = 5,
            AO = 6,
            AP = 7,
            AQ = 8,
            AR = 9,
            AS = 10,
            AT = 11,
            AU = 12,
            AV = 13,
            AW = 14,
            AX = 15,
            AY = 16,
            AZ = 17,
            BA = 18,
            BB = 19,
            BC = 20,
            BD = 21,
            BE = 22,
            BF = 23,
            BG = 24,
            BH = 25,
            BI = 26,
            BJ = 27,
            BK = 28,
            BL = 29,
            BM = 30,
            BN = 31,
            BO = 32,
            BP = 33,
            BQ = 34,
            BR = 35,
            BS = 36,
            BT = 37,
            BU = 38,
            BV = 39,
            BW = 40,
            BX = 41,
            BY = 42,
            BZ = 43,
            CA = 44,
            CB = 45,
            CC = 46,
            CD = 47,
            CE = 48,
            CF = 49,
            CG = 50,
            CH = 51,
            CI = 52,
            CJ = 53,
            CK = 54,
            CL = 55,
            CM = 56,
            CN = 57,
            CO = 58,
            CP = 59,
            CQ = 60,
            CR = 61,
            CS = 62,
            CT = 63,
            CU = 64,
            CV = 65,
            CW = 66,
            CX = 67,
            CY = 68,
            CZ = 69,
            DA = 70,
            DB = 71,
            DC = 72
        }

        /// <summary>
        /// Prices. Index corresponds with part size.
        /// </summary>
        internal enum Prices
        {
            DE = 1,
            DF = 2,
            DG = 3,
            DH = 4,
            DI = 5,
            DJ = 6,
            DK = 7,
            DL = 8,
            DM = 9,
            DN = 10,
            DO = 11,
            DP = 12,
            DQ = 13,
            DR = 14,
            DS = 15,
            DT = 16,
            DU = 17,
            DV = 18,
            DW = 19,
            DX = 20,
            DY = 21,
            DZ = 22,
            EA = 23,
            EB = 24,
            EC = 25,
            ED = 26,
            EE = 27,
            EF = 28,
            EG = 29,
            EH = 30,
            EI = 31,
            EJ = 32,
            EK = 33,
            EL = 34,
            EM = 35,
            EN = 36,
            EO = 37,
            EP = 38,
            EQ = 39,
            ER = 40,
            ES = 41,
            ET = 42,
            EU = 43,
            EV = 44,
            EW = 45,
            EX = 46,
            EY = 47,
            EZ = 48,
            FA = 49,
            FB = 50,
            FC = 51,
            FD = 52,
            FE = 53,
            FF = 54,
            FG = 55,
            FH = 56,
            FI = 57,
            FJ = 58,
            FK = 59,
            FL = 60,
            FM = 61,
            FN = 62,
            FO = 63,
            FP = 64,
            FQ = 65,
            FR = 66,
            FS = 67,
            FT = 68,
            FU = 69,
            FV = 70,
            FW = 71,
            FX = 72
        }

        internal const string ThreadTypeCheck = "Thread Type Check";
        internal const string ExternalThreadClass = "External Thread Class";

        internal enum Append
        {
            [Description(EpicorPartNumber)]
            A,
            [Description(WebSitePartNumber)]
            B,
            [Description(ThreadAngle)]
            C,
            [Description(ThreadTypeCheck)]
            D,
            [Description(ExternalThreadClass)]
            E,
            [Description(InternalThreadClass)]
            F,
            [Description(ThreadType)]
            G,
            [Description(ThreadClass)]
            H,
            [Description(ANominalDiameterInches)]
            I,
            [Description(ANominalDiameterMillimeters)]
            J,
            [Description(RRootDiameterMinimumInches)]
            K,
            [Description(RRootDiameterMinimumMillimeters)]
            L,
            [Description(DiameterCode)]
            M,
            [Description(ImperialOrMetric)]
            N,
            [Description(LeadInches)]
            O,
            [Description(LeadMillimeters)]
            P,
            [Description(LeadCode)]
            Q,
            [Description(PitchInches)]
            R,
            [Description(PitchMillimeters)]
            S,
            [Description(Starts)]
            T,
            [Description(TurnsPerInch)]
            U,
            [Description(ThreadsPerMillimeters)]
            V,
            [Description(Type123)]
            W,
            [Description(EndCodeForType4)]
            X,
            [Description(NutSize)]
            Y,
            [Description(HandRightOrLeft)]
            Z,
            [Description(ScrewMaterial)]
            AA,
            [Description(DiaUnit)]
            AB,
            [Description(LeadUnit)]
            AC,
            [Description(ColorCode)]
            AD,
            [Description(AcmeCode)]
            AE,
            [Description(ScrewWeight)]
            AF,
            [Description(WebSalable)]
            AG,
            [Description(LeadAccuracy)]
            AH,
            [Description(LeadTime)]
            AI,
            [Description(MarketingProductName)]
            FY,
            [Description(MarketingDescription)]
            FZ,
            [Description(Category1)]
            GA,
            [Description(Category2)]
            GB,
            [Description(Category3)]
            GC,
            [Description(Category4)]
            GD,
            [Description(Category5)]
            GE,
            [Description(CadLink)]
            GF,
            [Description(Document1)]
            GG,
            [Description(Document2)]
            GH,
            [Description(Document3)]
            GI,
            [Description(Document4)]
            GJ,
            [Description(Document5)]
            GK,
            [Description(VideoLink)]
            GL,
            [Description(Calculator)]
            GM,
            [Description(Image1)]
            GN,
            [Description(Image2)]
            GO,
            [Description(Image3)]
            GP,
            [Description(Image4)]
            GQ
        }
    }
}
