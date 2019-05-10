using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPLabSheetParserDLL.Resources;
using System.IO;
using System.Threading;
using System.Globalization;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;

namespace CSSPLabSheetParserDLL.Services
{
    public class CSSPLabSheetParser
    {
        #region Variables
        #endregion Variables

        #region Properties
        #endregion Properties

        #region Constructors
        public CSSPLabSheetParser()
        {
            // forcing all computers to run in English because all LabSheets decimal values are with a dot

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");
        }
        #endregion Constructors

        #region Functions public
        public string CheckFollowingAndCount(int LineNumber, string OldFirstObj, List<string> ValueArr, string ToFollow, int count)
        {
            string retStr = "";
            if (OldFirstObj != ToFollow)
            {
                retStr = string.Format(LabSheetParserRes.ErrorReadingFileAtLine_Error_, LineNumber, string.Format(LabSheetParserRes._HasToBeFollowing_InTheFile, ValueArr[0], ToFollow));
            }
            if (ValueArr.Count != count)
            {
                retStr = string.Format(LabSheetParserRes.ErrorReadingFileAtLine_Error_, LineNumber, string.Format(LabSheetParserRes._Requires_Value, ValueArr[0], count));
            }

            return retStr;
        }
        public string GetVariableValueStr(string variableStr, StringBuilder sb)
        {
            string retStr = "";

            string fileContent = sb.ToString();

            int varLength = ("[" + variableStr + "|").Length;
            int PosStart = fileContent.IndexOf("[" + variableStr + "|");
            if (PosStart == -1)
                return "ERR: Variable " + variableStr + " could not be found.";

            PosStart = PosStart + varLength;
            int PosEnd = fileContent.IndexOf("]", PosStart);

            if (PosEnd == -1)
                return "ERR: End ] could not be found for variable " + variableStr + ".";

            retStr = fileContent.Substring(PosStart, PosEnd - PosStart);

            return retStr.Trim();
        }
        public LabSheetA1Sheet ParseLabSheetA1(string LabSheetFileContent)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");

            StringBuilder sbPrevCommands = new StringBuilder();
            LabSheetA1Sheet labSheetA1Sheet = new LabSheetA1Sheet() { Error = "" };
            List<string> VarArr = new List<string>();
            StringBuilder sbFileContent = new StringBuilder(LabSheetFileContent);

            // Verison
            string retStr = GetVariableValueStr("Version", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.Version = int.Parse(retStr);

            // Sampling Plan Type
            retStr = GetVariableValueStr("Sampling Plan Type", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.SamplingPlanType = SamplingPlanTypeEnum.Error;
            for (int i = 1, count = Enum.GetNames(typeof(SamplingPlanTypeEnum)).Count(); i < count; i++)
            {
                if (((SamplingPlanTypeEnum)i).ToString() == retStr)
                {
                    labSheetA1Sheet.SamplingPlanType = (SamplingPlanTypeEnum)i;
                    break;
                }
            }

            // Sample Type
            retStr = GetVariableValueStr("Sample Type", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.SampleType = SampleTypeEnum.Error;
            for (int i = 101, count = Enum.GetNames(typeof(SampleTypeEnum)).Count() + 100; i < count; i++)
            {
                if (((SampleTypeEnum)i).ToString() == retStr)
                {
                    labSheetA1Sheet.SampleType = (SampleTypeEnum)i;
                    break;
                }
            }

            // Lab Sheet Type
            retStr = GetVariableValueStr("Lab Sheet Type", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.LabSheetType = LabSheetTypeEnum.Error;
            for (int i = 1, count = Enum.GetNames(typeof(LabSheetTypeEnum)).Count(); i < count; i++)
            {
                if (((LabSheetTypeEnum)i).ToString() == retStr)
                {
                    labSheetA1Sheet.LabSheetType = (LabSheetTypeEnum)i;
                    break;
                }
            }

            // Subsector
            retStr = GetVariableValueStr("Subsector", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            if (VarArr.Count != 2)
            {
                labSheetA1Sheet.Error = "ERR: Subsector variable should have 2 values";
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.SubsectorName = VarArr[0].Trim();
            labSheetA1Sheet.SubsectorTVItemID = int.Parse(VarArr[1].Trim());

            // Date
            retStr = GetVariableValueStr("Date", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            if (VarArr.Count != 3)
            {
                labSheetA1Sheet.Error = "ERR: Date variable should have 3 values";
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.RunYear = VarArr[0].Trim();
            labSheetA1Sheet.RunMonth = VarArr[1].Trim();
            labSheetA1Sheet.RunDay = VarArr[2].Trim();

            // Run
            retStr = GetVariableValueStr("Run", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            if (VarArr.Count != 1)
            {
                labSheetA1Sheet.Error = "ERR: Run variable should have 1 values";
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.RunNumber = int.Parse(VarArr[0].Trim());

            // Tides
            retStr = GetVariableValueStr("Tides", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.Tides = retStr;

            // Include Laboratory QA/QC
            retStr = GetVariableValueStr("IncludeLaboratoryQAQC", sbFileContent);

            VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

            if (VarArr.Count != 1)
            {
                labSheetA1Sheet.Error = "ERR: IncludeLaboratoryQAQC should have 1 values";
                return labSheetA1Sheet;
            }
            else
            {
                if (VarArr[0].ToLower().StartsWith("t"))
                {
                    labSheetA1Sheet.IncludeLaboratoryQAQC = true;
                }
                else
                {
                    labSheetA1Sheet.IncludeLaboratoryQAQC = false;
                }
            }

            if (labSheetA1Sheet.IncludeLaboratoryQAQC)
            {
                // Sample Crew Initials
                retStr = GetVariableValueStr("Sample Crew Initials", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.SampleCrewInitials = retStr;

                // Incubation Start Same Day
                retStr = GetVariableValueStr("Incubation Start Same Day", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.IncubationStartSameDay = retStr;

                // Incubation Start Time
                retStr = GetVariableValueStr("Incubation Start Time", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Incubation Start Time variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.IncubationBath1StartTime = VarArr[0].Trim();
                labSheetA1Sheet.IncubationBath2StartTime = VarArr[1].Trim();
                labSheetA1Sheet.IncubationBath3StartTime = VarArr[2].Trim();

                // Incubation End Time
                retStr = GetVariableValueStr("Incubation End Time", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Incubation End Time variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.IncubationBath1EndTime = VarArr[0].Trim();
                labSheetA1Sheet.IncubationBath2EndTime = VarArr[1].Trim();
                labSheetA1Sheet.IncubationBath3EndTime = VarArr[2].Trim();

                // Incubation Time Calculated
                retStr = GetVariableValueStr("Incubation Time Calculated", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Incubation Time Calculated variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.IncubationBath1TimeCalculated = VarArr[0].Trim();
                labSheetA1Sheet.IncubationBath2TimeCalculated = VarArr[1].Trim();
                labSheetA1Sheet.IncubationBath3TimeCalculated = VarArr[2].Trim();

                // Water Bath Count
                retStr = GetVariableValueStr("Water Bath Count", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.WaterBathCount = int.Parse(retStr);

                // Water Bath
                retStr = GetVariableValueStr("Water Bath", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Water Bath variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.WaterBath1 = VarArr[0].Trim();
                labSheetA1Sheet.WaterBath2 = VarArr[1].Trim();
                labSheetA1Sheet.WaterBath3 = VarArr[2].Trim();

                // TC Has 2 Coolers
                retStr = GetVariableValueStr("TC Has 2 Coolers", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.TCHas2Coolers = retStr;

                // TC Field
                retStr = GetVariableValueStr("TC Field", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 2)
                {
                    labSheetA1Sheet.Error = "ERR: TC Field variable should have 2 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.TCField1 = VarArr[0].Trim();
                labSheetA1Sheet.TCField2 = VarArr[1].Trim();

                // TC Lab
                retStr = GetVariableValueStr("TC Lab", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 2)
                {
                    labSheetA1Sheet.Error = "ERR: TC Lab variable should have 2 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.TCLab1 = VarArr[0].Trim();
                labSheetA1Sheet.TCLab2 = VarArr[1].Trim();

                // TC First
                retStr = GetVariableValueStr("TC First", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.TCFirst = retStr;

                // TC Average
                retStr = GetVariableValueStr("TC Average", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.TCAverage = retStr;

                // Control Lot
                retStr = GetVariableValueStr("Control Lot", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.ControlLot = retStr;

                // Positive 35
                retStr = GetVariableValueStr("Positive 35", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.Positive35 = retStr;

                // Non Target 35
                retStr = GetVariableValueStr("Non Target 35", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.NonTarget35 = retStr;

                // Negative 35
                retStr = GetVariableValueStr("Negative 35", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.Negative35 = retStr;

                // Positive 44.5
                retStr = GetVariableValueStr("Positive 44.5", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Positive 44.5 variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.Bath1Positive44_5 = VarArr[0].Trim();
                labSheetA1Sheet.Bath2Positive44_5 = VarArr[1].Trim();
                labSheetA1Sheet.Bath3Positive44_5 = VarArr[2].Trim();

                // Non Target 44.5
                retStr = GetVariableValueStr("Non Target 44.5", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Non Target 44.5 variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.Bath1NonTarget44_5 = VarArr[0].Trim();
                labSheetA1Sheet.Bath2NonTarget44_5 = VarArr[1].Trim();
                labSheetA1Sheet.Bath3NonTarget44_5 = VarArr[2].Trim();

                // Negative 44.5
                retStr = GetVariableValueStr("Negative 44.5", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Negative 44.5 variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.Bath1Negative44_5 = VarArr[0].Trim();
                labSheetA1Sheet.Bath2Negative44_5 = VarArr[1].Trim();
                labSheetA1Sheet.Bath3Negative44_5 = VarArr[2].Trim();

                // Blank 35
                retStr = GetVariableValueStr("Blank 35", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.Blank35 = retStr;

                // Blank 44.5
                retStr = GetVariableValueStr("Blank 44.5", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Blank 44.5 variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.Bath1Blank44_5 = VarArr[0].Trim();
                labSheetA1Sheet.Bath2Blank44_5 = VarArr[1].Trim();
                labSheetA1Sheet.Bath3Blank44_5 = VarArr[2].Trim();

                // Lot 35
                retStr = GetVariableValueStr("Lot 35", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.Lot35 = retStr;

                // Lot 44.5
                retStr = GetVariableValueStr("Lot 44.5", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                    labSheetA1Sheet.Error = retStr;

                labSheetA1Sheet.Lot44_5 = retStr;

                // Daily Duplicate
                retStr = GetVariableValueStr("Daily Duplicate", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Daily Duplicate variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.DailyDuplicateRLog = VarArr[0].Trim();
                labSheetA1Sheet.DailyDuplicatePrecisionCriteria = VarArr[1].Trim();
                labSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable = VarArr[2].Trim();

                // Intertech Duplicate
                retStr = GetVariableValueStr("Intertech Duplicate", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Intertech Duplicate variable should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.IntertechDuplicateRLog = VarArr[0].Trim();
                labSheetA1Sheet.IntertechDuplicatePrecisionCriteria = VarArr[1].Trim();
                labSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable = VarArr[2].Trim();

                // Intertech Read
                retStr = GetVariableValueStr("Intertech Read", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.IntertechReadAcceptableOrUnacceptable = retStr;

                // Sample Bottle Lot Number
                retStr = GetVariableValueStr("Sample Bottle Lot Number", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.SampleBottleLotNumber = retStr;

                // Salinities
                retStr = GetVariableValueStr("Salinities", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 4)
                {
                    labSheetA1Sheet.Error = "ERR: Salinities variable should have 4 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.SalinitiesReadBy = VarArr[0].Trim();
                labSheetA1Sheet.SalinitiesReadYear = VarArr[1].Trim();
                labSheetA1Sheet.SalinitiesReadMonth = VarArr[2].Trim();
                labSheetA1Sheet.SalinitiesReadDay = VarArr[3].Trim();

                // Results
                retStr = GetVariableValueStr("Results", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 4)
                {
                    labSheetA1Sheet.Error = "ERR: Results variable should have 4 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.ResultsReadBy = VarArr[0].Trim();
                labSheetA1Sheet.ResultsReadYear = VarArr[1].Trim();
                labSheetA1Sheet.ResultsReadMonth = VarArr[2].Trim();
                labSheetA1Sheet.ResultsReadDay = VarArr[3].Trim();

                // Recorded
                retStr = GetVariableValueStr("Recorded", sbFileContent);
                if (retStr.StartsWith("ERR:"))
                {
                    labSheetA1Sheet.Error = retStr;
                    return labSheetA1Sheet;
                }

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 4)
                {
                    labSheetA1Sheet.Error = "ERR: Recorded variable should have 4 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.ResultsRecordedBy = VarArr[0].Trim();
                labSheetA1Sheet.ResultsRecordedYear = VarArr[1].Trim();
                labSheetA1Sheet.ResultsRecordedMonth = VarArr[2].Trim();
                labSheetA1Sheet.ResultsRecordedDay = VarArr[3].Trim();




                // Approved By Supervisor Initials
                retStr = GetVariableValueStr("Approved By Supervisor Initials", sbFileContent);

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 1)
                {
                    labSheetA1Sheet.Error = "ERR: Approved By Supervisor Initials should have 1 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.ApprovedBySupervisorInitials = VarArr[0].Trim();

                // Approval Date
                retStr = GetVariableValueStr("Approval Date", sbFileContent);

                VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                if (VarArr.Count != 3)
                {
                    labSheetA1Sheet.Error = "ERR: Approval Date should have 3 values";
                    return labSheetA1Sheet;
                }

                labSheetA1Sheet.ApprovalYear = VarArr[0].Trim();
                labSheetA1Sheet.ApprovalMonth = VarArr[1].Trim();
                labSheetA1Sheet.ApprovalDay = VarArr[2].Trim();

            }
            else
            {
                if (labSheetA1Sheet.Version == 2)
                {
                    // Daily Duplicate
                    retStr = GetVariableValueStr("Daily Duplicate", sbFileContent);
                    if (retStr.StartsWith("ERR:"))
                    {
                        labSheetA1Sheet.Error = retStr;
                        return labSheetA1Sheet;
                    }

                    VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                    if (VarArr.Count != 3)
                    {
                        labSheetA1Sheet.Error = "ERR: Daily Duplicate variable should have 3 values";
                        return labSheetA1Sheet;
                    }

                    labSheetA1Sheet.DailyDuplicateRLog = VarArr[0].Trim();
                    labSheetA1Sheet.DailyDuplicatePrecisionCriteria = VarArr[1].Trim();
                    labSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable = VarArr[2].Trim();

                    // Intertech Duplicate
                    retStr = GetVariableValueStr("Intertech Duplicate", sbFileContent);
                    if (retStr.StartsWith("ERR:"))
                    {
                        labSheetA1Sheet.Error = retStr;
                        return labSheetA1Sheet;
                    }

                    VarArr = retStr.Split("|".ToCharArray(), StringSplitOptions.None).ToList();

                    if (VarArr.Count != 3)
                    {
                        labSheetA1Sheet.Error = "ERR: Intertech Duplicate variable should have 3 values";
                        return labSheetA1Sheet;
                    }

                    labSheetA1Sheet.IntertechDuplicateRLog = VarArr[0].Trim();
                    labSheetA1Sheet.IntertechDuplicatePrecisionCriteria = VarArr[1].Trim();
                    labSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable = VarArr[2].Trim();

                    // Intertech Read
                    retStr = GetVariableValueStr("Intertech Read", sbFileContent);
                    if (retStr.StartsWith("ERR:"))
                    {
                        labSheetA1Sheet.Error = retStr;
                        return labSheetA1Sheet;
                    }

                    labSheetA1Sheet.IntertechReadAcceptableOrUnacceptable = retStr;
                }
            }

            // Run Weather Comment
            retStr = GetVariableValueStr("Run Weather Comment", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.RunWeatherComment = retStr;

            // Run Comment
            retStr = GetVariableValueStr("Run Comment", sbFileContent);
            if (retStr.StartsWith("ERR:"))
            {
                labSheetA1Sheet.Error = retStr;
                return labSheetA1Sheet;
            }

            labSheetA1Sheet.RunComment = retStr;

            using (StringReader sr = new StringReader(LabSheetFileContent))
            {
                int LineNumber = 0;
                string OldFirstObj = "";
                string lineStr = "";
                while ((lineStr = sr.ReadLine()) != null)
                {
                    if (!(lineStr.StartsWith("Site") || lineStr.StartsWith("Log")))
                        continue;

                    LineNumber += 1;
                    List<string> ValueArr = lineStr.Split("\t".ToCharArray(), StringSplitOptions.None).ToList();

                    for (int i = 0, count = ValueArr.Count; i < count; i++)
                    {
                        ValueArr[i] = ValueArr[i].Trim();
                    }

                    switch (ValueArr[0])
                    {
                        case "Site":
                            {
                                while ((lineStr = sr.ReadLine()) != null)
                                {
                                    if (lineStr == "________________________________")
                                        break;

                                    ValueArr = lineStr.Split("\t".ToCharArray(), StringSplitOptions.None).ToList();

                                    for (int i = 0, count = ValueArr.Count; i < count; i++)
                                    {
                                        ValueArr[i] = ValueArr[i].Trim();
                                    }

                                    if (ValueArr.Count != 13)
                                    {
                                        labSheetA1Sheet.Error = string.Format(LabSheetParserRes.ErrorReadingFileAtLine_Error_, LineNumber, string.Format(LabSheetParserRes._Requires_Value, ValueArr[0], 13));
                                        return labSheetA1Sheet;
                                    }

                                    List<string> ValueList = new List<string>();
                                    int col = 0;
                                    LabSheetA1Measurement a1Measurement = new LabSheetA1Measurement();
                                    foreach (string s in ValueArr)
                                    {
                                        ValueList.Add(s);
                                        switch (col)
                                        {
                                            case 0: // Site
                                                {
                                                    if (string.IsNullOrWhiteSpace(s))
                                                    {
                                                        labSheetA1Sheet.Error = string.Format(LabSheetParserRes.ErrorReadingFileAtLine_Error_, LineNumber, string.Format(LabSheetParserRes._IsRequired, "Site"));
                                                        return labSheetA1Sheet;
                                                    }
                                                    a1Measurement.Site = s;
                                                }
                                                break;
                                            case 1: // Time converted into date for DB
                                                {
                                                    a1Measurement.Time = null;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        int Year = 0;
                                                        int Month = 0;
                                                        int Day = 0;
                                                        int Hour = 0;
                                                        int Minute = 0;

                                                        int.TryParse(labSheetA1Sheet.RunYear, out Year);
                                                        int.TryParse(labSheetA1Sheet.RunMonth, out Month);
                                                        int.TryParse(labSheetA1Sheet.RunDay, out Day);

                                                        if (s.Length == 5)
                                                        {
                                                            int.TryParse(s.Substring(0, 2), out Hour);
                                                            int.TryParse(s.Substring(3, 2), out Minute);
                                                        }

                                                        if (Hour > 23)
                                                            Hour = 0;

                                                        if (Minute > 59)
                                                            Minute = 0;

                                                        a1Measurement.Time = new DateTime(Year, Month, Day, Hour, Minute, 0);
                                                    }
                                                }
                                                break;
                                            case 2: // MPN
                                                {
                                                    a1Measurement.MPN = null;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        int TempInt = -1;
                                                        if (s.Trim().StartsWith("<"))
                                                        {
                                                            TempInt = 1;
                                                        }
                                                        else if (s.Trim().StartsWith(">"))
                                                        {
                                                            TempInt = 1700;
                                                        }
                                                        else
                                                        {
                                                            if (!int.TryParse(s, out TempInt))
                                                            {
                                                                TempInt = -999;
                                                            }
                                                        }
                                                        a1Measurement.MPN = TempInt;
                                                    }
                                                }
                                                break;
                                            case 3: // Tube 10
                                                {
                                                    a1Measurement.Tube10 = null;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        int TempInt = -1;
                                                        int.TryParse(s, out TempInt);
                                                        if (TempInt == -1)
                                                            continue;

                                                        if (TempInt >= 0)
                                                            a1Measurement.Tube10 = TempInt;
                                                    }
                                                }
                                                break;
                                            case 4: // Tube 1.0
                                                {
                                                    a1Measurement.Tube1_0 = null;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        int TempInt = -1;
                                                        int.TryParse(s, out TempInt);
                                                        if (TempInt == -1)
                                                            continue;

                                                        if (TempInt >= 0)
                                                            a1Measurement.Tube1_0 = TempInt;
                                                    }
                                                }
                                                break;
                                            case 5: // Tube 0.1
                                                {
                                                    a1Measurement.Tube0_1 = null;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        int TempInt = -1;
                                                        int.TryParse(s, out TempInt);
                                                        if (TempInt == -1)
                                                            continue;

                                                        if (TempInt >= 0)
                                                            a1Measurement.Tube0_1 = TempInt;
                                                    }
                                                }
                                                break;
                                            case 6: // Salinity
                                                {
                                                    a1Measurement.Salinity = null;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        float TempFloat = -1;
                                                        float.TryParse(s, out TempFloat);
                                                        if (TempFloat == -1)
                                                            continue;

                                                        if (TempFloat >= 0)
                                                            a1Measurement.Salinity = TempFloat;
                                                    }
                                                }
                                                break;
                                            case 7: // Temperature
                                                {
                                                    a1Measurement.Temperature = null;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        float TempFloat = -99.0f;
                                                        float.TryParse(s, out TempFloat);
                                                        if (TempFloat == -99.0f)
                                                            continue;

                                                        a1Measurement.Temperature = TempFloat;
                                                    }
                                                }
                                                break;
                                            case 8: // ProcessedBy
                                                {
                                                    a1Measurement.ProcessedBy = null;

                                                    if (!string.IsNullOrWhiteSpace(s))
                                                        a1Measurement.ProcessedBy = s;
                                                }
                                                break;
                                            case 9: // SampleType
                                                {
                                                    a1Measurement.SampleType = SampleTypeEnum.Error;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        for (int i = 1, count = Enum.GetNames(typeof(SampleTypeEnum)).Length; i < count; i++)
                                                        {
                                                            int j = i + 100;
                                                            string SampleTypeTest = ((SampleTypeEnum)j).ToString();
                                                            SampleTypeEnum SampleTypeValue = (SampleTypeEnum)j;
                                                            if (SampleTypeTest == s)
                                                            {
                                                                a1Measurement.SampleType = SampleTypeValue;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                                break;
                                            case 10: // TVItemID of the MWQMSite
                                                {
                                                    a1Measurement.TVItemID = 0;
                                                    if (!string.IsNullOrWhiteSpace(s))
                                                    {
                                                        int TempInt = 0;
                                                        if (!int.TryParse(s, out TempInt))
                                                        {
                                                            TempInt = 0;
                                                        }
                                                        a1Measurement.TVItemID = TempInt;
                                                    }
                                                }
                                                break;
                                            case 11: // SiteComment
                                                {
                                                    a1Measurement.SiteComment = null;

                                                    if (!string.IsNullOrWhiteSpace(s))
                                                        a1Measurement.SiteComment = s;
                                                }
                                                break;
                                            default:
                                                break;
                                        }
                                        col += 1;
                                    }

                                    if (!string.IsNullOrWhiteSpace(a1Measurement.Site))
                                    {
                                        labSheetA1Sheet.LabSheetA1MeasurementList.Add(a1Measurement);
                                    }
                                }

                            }
                            break;
                        case "Log":
                            {
                                sbPrevCommands.AppendLine("________________________________");
                                sbPrevCommands.AppendLine("Log");
                                while ((lineStr = sr.ReadLine()) != null)
                                {
                                    sbPrevCommands.AppendLine(lineStr);
                                }
                                labSheetA1Sheet.Log = sbPrevCommands.ToString();
                            }
                            break;
                        default:
                            {
                                labSheetA1Sheet.Error = string.Format(LabSheetParserRes.ErrorReadingFileAtLine_Error_, LineNumber, string.Format(LabSheetParserRes.UnknownParameter_, ValueArr[0]));
                                return labSheetA1Sheet;
                            }
                    }

                    OldFirstObj = ValueArr[0];
                }
            }

            return labSheetA1Sheet;
        }
        #endregion Functions public
    }
}