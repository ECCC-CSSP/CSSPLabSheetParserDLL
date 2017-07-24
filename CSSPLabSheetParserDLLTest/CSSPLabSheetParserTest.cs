using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using CSSPLabSheetParserDLL.Services;
using CSSPLabSheetParserDLL.Resources;
using System.Threading;
using System.Globalization;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;

namespace CSSPLabSheetParserDLLTest
{
    /// <summary>
    /// Summary description for CSSPLabSheetParserTest
    /// </summary>
    [TestClass]
    public class CSSPLabSheetParserTest
    {
        #region Variables
        string TestFileName = @"C:\CSSP latest code\CSSPLabSheetParserDLL\CSSPLabSheetParserDLLTest\LabSheetTestFile.txt";
        private TestContext testContextInstance;
        #endregion Variables

        #region Properties
        public CSSPLabSheetParser csspLabSheetParser { get; set; }
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        #endregion Properties

        #region Constructors
        public CSSPLabSheetParserTest()
        {
        }
        #endregion Constructors

        #region Initialize and Cleanup
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion Initialize and Cleanup

        #region Testing functions
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Good_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.AreEqual("", labSheetA1Sheet.Error);
            Assert.AreEqual(1, labSheetA1Sheet.Version);
            Assert.AreEqual(SamplingPlanTypeEnum.Subsector, labSheetA1Sheet.SamplingPlanType);
            Assert.AreEqual(SampleTypeEnum.Routine, labSheetA1Sheet.SampleType);
            Assert.AreEqual(LabSheetTypeEnum.A1, labSheetA1Sheet.LabSheetType);
            Assert.AreEqual("NB-01-020-002 (Charlo)", labSheetA1Sheet.SubsectorName);
            Assert.AreEqual(560, labSheetA1Sheet.SubsectorTVItemID);
            Assert.AreEqual("2016", labSheetA1Sheet.RunYear);
            Assert.AreEqual("10", labSheetA1Sheet.RunMonth);
            Assert.AreEqual("25", labSheetA1Sheet.RunDay);
            Assert.AreEqual("HT / HF", labSheetA1Sheet.Tides);
            Assert.AreEqual("KJ,KDF", labSheetA1Sheet.SampleCrewInitials);
            Assert.AreEqual("true", labSheetA1Sheet.IncubationStartSameDay);
            Assert.AreEqual(3, labSheetA1Sheet.WaterBathCount);
            Assert.AreEqual("13:13", labSheetA1Sheet.IncubationBath1StartTime);
            Assert.AreEqual("12:12", labSheetA1Sheet.IncubationBath2StartTime);
            Assert.AreEqual("14:14", labSheetA1Sheet.IncubationBath3StartTime);
            Assert.AreEqual("14:12", labSheetA1Sheet.IncubationBath1EndTime);
            Assert.AreEqual("12:34", labSheetA1Sheet.IncubationBath2EndTime);
            Assert.AreEqual("14:23", labSheetA1Sheet.IncubationBath3EndTime);
            Assert.AreEqual("24:59", labSheetA1Sheet.IncubationBath1TimeCalculated);
            Assert.AreEqual("24:22", labSheetA1Sheet.IncubationBath2TimeCalculated);
            Assert.AreEqual("24:09", labSheetA1Sheet.IncubationBath3TimeCalculated);
            Assert.AreEqual("H", labSheetA1Sheet.WaterBath1);
            Assert.AreEqual("G", labSheetA1Sheet.WaterBath2);
            Assert.AreEqual("F", labSheetA1Sheet.WaterBath3);
            Assert.AreEqual("true", labSheetA1Sheet.TCHas2Coolers);
            Assert.AreEqual("22", labSheetA1Sheet.TCField1);
            Assert.AreEqual("23", labSheetA1Sheet.TCField2);
            Assert.AreEqual("7", labSheetA1Sheet.TCLab1);
            Assert.AreEqual("8", labSheetA1Sheet.TCLab2);
            Assert.AreEqual("3.3", labSheetA1Sheet.TCFirst);
            Assert.AreEqual("2.4", labSheetA1Sheet.TCAverage);
            Assert.AreEqual("Something from lot", labSheetA1Sheet.ControlLot);
            Assert.AreEqual("+", labSheetA1Sheet.Positive35);
            Assert.AreEqual("+", labSheetA1Sheet.NonTarget35);
            Assert.AreEqual("-", labSheetA1Sheet.Negative35);
            Assert.AreEqual("+", labSheetA1Sheet.Bath1Positive44_5);
            Assert.AreEqual("+", labSheetA1Sheet.Bath2Positive44_5);
            Assert.AreEqual("+", labSheetA1Sheet.Bath3Positive44_5);
            Assert.AreEqual("-", labSheetA1Sheet.Bath1NonTarget44_5);
            Assert.AreEqual("-", labSheetA1Sheet.Bath2NonTarget44_5);
            Assert.AreEqual("-", labSheetA1Sheet.Bath3NonTarget44_5);
            Assert.AreEqual("-", labSheetA1Sheet.Bath1Negative44_5);
            Assert.AreEqual("-", labSheetA1Sheet.Bath2Negative44_5);
            Assert.AreEqual("-", labSheetA1Sheet.Bath3Negative44_5);
            Assert.AreEqual("-", labSheetA1Sheet.Blank35);
            Assert.AreEqual("-", labSheetA1Sheet.Bath1Blank44_5);
            Assert.AreEqual("+", labSheetA1Sheet.Bath2Blank44_5);
            Assert.AreEqual("-", labSheetA1Sheet.Bath3Blank44_5);
            Assert.AreEqual("87", labSheetA1Sheet.Lot35);
            Assert.AreEqual("85", labSheetA1Sheet.Lot44_5);
            Assert.AreEqual("3.23045", labSheetA1Sheet.DailyDuplicateRLog);
            Assert.AreEqual("0.6872", labSheetA1Sheet.DailyDuplicatePrecisionCriteria);
            Assert.AreEqual("Unacceptable", labSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable);
            Assert.AreEqual("1.23045", labSheetA1Sheet.IntertechDuplicateRLog);
            Assert.AreEqual("0.093", labSheetA1Sheet.IntertechDuplicatePrecisionCriteria);
            Assert.AreEqual("Unacceptable", labSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable);
            Assert.AreEqual("Unacceptable", labSheetA1Sheet.IntertechReadAcceptableOrUnacceptable);
            Assert.AreEqual("This is the weather comment", labSheetA1Sheet.RunWeatherComment);
            Assert.AreEqual("This is the Run comment", labSheetA1Sheet.RunComment);
            Assert.AreEqual("87,45", labSheetA1Sheet.SampleBottleLotNumber);
            Assert.AreEqual("KJ", labSheetA1Sheet.SalinitiesReadBy);
            Assert.AreEqual("2016", labSheetA1Sheet.SalinitiesReadYear);
            Assert.AreEqual("10", labSheetA1Sheet.SalinitiesReadMonth);
            Assert.AreEqual("25", labSheetA1Sheet.SalinitiesReadDay);
            Assert.AreEqual("JH", labSheetA1Sheet.ResultsReadBy);
            Assert.AreEqual("2016", labSheetA1Sheet.ResultsReadYear);
            Assert.AreEqual("10", labSheetA1Sheet.ResultsReadMonth);
            Assert.AreEqual("26", labSheetA1Sheet.ResultsReadDay);
            Assert.AreEqual("HG", labSheetA1Sheet.ResultsRecordedBy);
            Assert.AreEqual("2016", labSheetA1Sheet.ResultsRecordedYear);
            Assert.AreEqual("10", labSheetA1Sheet.ResultsRecordedMonth);
            Assert.AreEqual("26", labSheetA1Sheet.ResultsRecordedDay);
            Assert.AreEqual("0022", labSheetA1Sheet.LabSheetA1MeasurementList[0].Site);
            Assert.AreEqual("13:13", labSheetA1Sheet.LabSheetA1MeasurementList[0].Time.Value.ToString("HH:mm"));
            Assert.AreEqual(4, labSheetA1Sheet.LabSheetA1MeasurementList[0].Tube10);
            Assert.AreEqual(2, labSheetA1Sheet.LabSheetA1MeasurementList[0].Tube1_0);
            Assert.AreEqual(2, labSheetA1Sheet.LabSheetA1MeasurementList[0].Tube0_1);
            Assert.AreEqual(1.0f, labSheetA1Sheet.LabSheetA1MeasurementList[0].Salinity);
            Assert.AreEqual(2.0f, labSheetA1Sheet.LabSheetA1MeasurementList[0].Temperature);
            Assert.AreEqual("ER", labSheetA1Sheet.LabSheetA1MeasurementList[0].ProcessedBy);
            Assert.AreEqual(SampleTypeEnum.Routine, labSheetA1Sheet.LabSheetA1MeasurementList[0].SampleType);
            Assert.AreEqual(32, labSheetA1Sheet.LabSheetA1MeasurementList[0].MPN);
            Assert.AreEqual(7153, labSheetA1Sheet.LabSheetA1MeasurementList[0].TVItemID);
            Assert.AreEqual("This is a comment for site 22", labSheetA1Sheet.LabSheetA1MeasurementList[0].SiteComment);

            Assert.AreEqual(1, labSheetA1Sheet.LabSheetA1MeasurementList[1].MPN);

            Assert.AreEqual(6, labSheetA1Sheet.LabSheetA1MeasurementList[2].MPN);

            Assert.AreEqual(SampleTypeEnum.Routine, labSheetA1Sheet.LabSheetA1MeasurementList[3].SampleType);
            Assert.AreEqual(-999, labSheetA1Sheet.LabSheetA1MeasurementList[3].MPN);

            Assert.AreEqual(SampleTypeEnum.DailyDuplicate, labSheetA1Sheet.LabSheetA1MeasurementList[4].SampleType);
            Assert.AreEqual(SampleTypeEnum.IntertechDuplicate, labSheetA1Sheet.LabSheetA1MeasurementList[5].SampleType);
            Assert.AreEqual(SampleTypeEnum.IntertechRead, labSheetA1Sheet.LabSheetA1MeasurementList[6].SampleType);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Empty_First_Line_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            List<string> variableStrList = new List<string>()
            {
                 "Version", "Sampling Plan Type", "Sample Type", "Lab Sheet Type", "Subsector", "Date", "Tides", "Sample Crew Initials",
                 "Incubation Start Same Day", "Incubation Start Time", "Incubation End Time", "Incubation Time Calculated", "Water Bath Count",
                 "Water Bath", "TC Has 2 Coolers","TC Field", "TC Lab", "TC First", "TC Average", "Control Lot", "Positive 35", "Non Target 35",
                 "Negative 35", "Positive 44.5", "Non Target 44.5", "Negative 44.5", "Blank 35", "Blank 44.5", "Lot 35", "Lot 44.5",
                 "Daily Duplicate", "Intertech Duplicate", "Intertech Read", "Run Weather Comment", "Run Comment", "Sample Bottle Lot Number",
                 "Salinities", "Results", "Recorded"
            };

            foreach (string variableStr in variableStrList)
            {
                string FullFileText2 = FullFileText.Replace((variableStr == "Water Bath" ? "Water Bath|" : variableStr), ""); ;

                LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText2);
                Assert.AreEqual("ERR: Variable " + variableStr + " could not be found.", labSheetA1Sheet.Error);
            }
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Site_Parameter_requires_X_values_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr + "\ta\ta\ta";
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.AreEqual(string.Format(LabSheetParserRes.ErrorReadingFileAtLine_Error_, 1, string.Format(LabSheetParserRes._Requires_Value, "0022", 13)), labSheetA1Sheet.Error);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Site_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            string VarText = "Site";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace("0022", "");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.AreEqual(string.Format(LabSheetParserRes.ErrorReadingFileAtLine_Error_, 1, string.Format(LabSheetParserRes._IsRequired, VarText)), labSheetA1Sheet.Error);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Time_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace("0022   	13:13", "0022   	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.IsNull(labSheetA1Sheet.LabSheetA1MeasurementList[0].Time);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Tube10_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace("0022   	13:13 	4", "0022   	13:13 	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.IsNull(labSheetA1Sheet.LabSheetA1MeasurementList[0].Tube10);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Tube1_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace("0022   	13:13 	4      	2", "0022   	13:13 	4      	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.IsNull(labSheetA1Sheet.LabSheetA1MeasurementList[0].Tube1_0);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Tube01_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace("0022   	13:13 	4      	2      	2", "0022   	13:13 	4      	2      	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.IsNull(labSheetA1Sheet.LabSheetA1MeasurementList[0].Tube0_1);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Salinity_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace("0022   	13:13 	4      	2      	2      	1.0", "0022   	13:13 	4      	2      	2      	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.IsNull(labSheetA1Sheet.LabSheetA1MeasurementList[0].Salinity);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_Temperature_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace("0022   	13:13 	4      	2      	2      	1.0  	2.0", "0022   	13:13 	4      	2      	2      	1.0  	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.IsNull(labSheetA1Sheet.LabSheetA1MeasurementList[0].Temperature);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_ProcessBy_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace("0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER",
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.IsNull(labSheetA1Sheet.LabSheetA1MeasurementList[0].ProcessedBy);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_SampleType_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace(
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER      	Routine",
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER      	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.AreEqual(SampleTypeEnum.Error, labSheetA1Sheet.LabSheetA1MeasurementList[0].SampleType);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_MPN_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace(
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER      	Routine       	32",
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER      	Routine       	aa");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.AreEqual("", labSheetA1Sheet.Error);
            Assert.AreEqual(-999, labSheetA1Sheet.LabSheetA1MeasurementList[0].MPN);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_TVItemID_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace(
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER      	Routine       	32        	7153",
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER      	Routine       	32        	aa");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.AreEqual(0, labSheetA1Sheet.LabSheetA1MeasurementList[0].TVItemID);
        }
        [TestMethod]
        public void LabSheetService_ParseLabSheetA1_SiteComment_Parameter_Empty_Error_Test()
        {
            Setup();

            FileInfo fiLabSheetTestFile = new FileInfo(TestFileName);
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            string FullFileContent = FullFileText;
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            StringReader sr2 = new StringReader(FullFileContent);
            int LineNumber = 0;
            string StartText = "0022   	13:13";
            for (int j = 0; j < 10000; j++)
            {
                string LineStr = sr2.ReadLine();

                if (LineStr == null)
                    break;

                if (LineStr.StartsWith(StartText) && LineNumber == 0)
                {
                    LineNumber = j;
                    LineStr = LineStr.Replace(
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER      	Routine       	32        	7153    	This is a comment for site 22",
                        "0022   	13:13 	4      	2      	2      	1.0  	2.0  	ER      	Routine       	32        	7153    	");
                    sb.AppendLine(LineStr);
                }
                else
                {
                    sb.AppendLine(LineStr);
                }
            }
            FullFileText = sb.ToString();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.IsNull(labSheetA1Sheet.LabSheetA1MeasurementList[0].SiteComment);
        }
        #endregion Testing functions

        #region Functions private
        public void Setup()
        {
            csspLabSheetParser = new CSSPLabSheetParser();
        }
        #endregion Functions private
    }
}
