using OfficeOpenXml;
using System;
using System.IO;
using System.Xml.Serialization;
using XMLToExcellConverter.Messages;
using XMLToExcellConverter.XmlObjects;

namespace XMLToExcellConverter
{


    class Program
    {
        static void Main(string[] args)
        {
            ShowStartOptions();
        }

        private static void ShowStartOptions()
        {
            const ConsoleKey quitKey = ConsoleKey.Escape;
            ConsoleKeyInfo pressedKey;

            do
            {
                Console.WriteLine("XML to XLS Converter");
                Console.WriteLine("Select Action:");
                Console.WriteLine("[1] - Create a xls from a sample file with fixed fields (Id,IsElite,GoldCost,Exp)");
                WriteInColor("[2] - Create a xls from any XML file - Disabled", ConsoleColor.DarkGray);
                Console.WriteLine("[ESC] - Exit");
                pressedKey = Console.ReadKey();

                switch (pressedKey.Key)
                {
                    case ConsoleKey.D1:
                        RunWithErrorHandling(() => ConvertSampleXMLToExcel());
                        break;
                    case ConsoleKey.D2:
                        Console.Clear();
                        WriteInColor("Disabled option, please type a valid option", ConsoleColor.Red);
                        Console.ReadKey();
                        Console.Clear();
                        break;
                    case quitKey:
                        Environment.Exit(0);
                        break;
                    default:
                        Console.Clear();
                        WriteInColor("Unknown option, please type a valid option", ConsoleColor.Red);
                        Console.ReadKey();
                        Console.Clear();
                        break;
                }
            }
            while (pressedKey.Key != quitKey);
        }

        private static void ConvertSampleXMLToExcel()
        {
            var awards = DesserializeSampleXML();
            var response = CreateExcellFileWithSampleXML(awards);

            if (response.Success)
            {
                Console.WriteLine();
                WriteInColor("Sample XLS File successfully generated. Opening it (sample.xlsx)...", ConsoleColor.Green);
                WriteInColor("Press Any key to quit", ConsoleColor.Green);
                System.Diagnostics.Process.Start(response.FileName);
                Console.ReadKey();
                Environment.Exit(0);
            }
            else
            {
                WriteInColor("There was an error generating Excel file", ConsoleColor.Red);
            }
        }


        #region Manipulating Sample XML
        private static AwardProps DesserializeSampleXML()
        {
            string filename = "sampleFile.xml";

            AwardProps awards;
            using (var reader = new StreamReader(filename))
            {
                XmlSerializer deserializer = new XmlSerializer(typeof(AwardProps), new XmlRootAttribute("AwardProps"));
                awards = (AwardProps)deserializer.Deserialize(reader);
            }

            return awards;
        }

        private static FileNameOperationResponse CreateExcellFileWithSampleXML(AwardProps awards)
        {
            string newFile = "sample.xlsx";

            if (File.Exists(newFile))
                File.Delete(newFile);

            var fileInfo = new FileInfo(newFile);

            using (ExcelPackage pck = new ExcelPackage(fileInfo))
            {
                var ws = pck.Workbook.Worksheets.Add("Content");
                ws.View.ShowGridLines = false;

                int awardCount = awards.AwardPropRecord.G_AwardProps.Entry.Count;

                ws.Cells["A1:D1"].Style.Font.Bold = true;
                ws.Cells["A1"].Value = "Id";
                ws.Cells["B1"].Value = "IsElite";
                ws.Cells["C1"].Value = "GoldCost";
                ws.Cells["D1"].Value = "Exp";

                for (int i = 0; i < awardCount; i++)
                {
                    var currentAward = awards.AwardPropRecord.G_AwardProps.Entry[i];

                    int InitialLine = 2;
                    int printLine = (InitialLine + i);

                    ws.Cells["A" + (printLine)].Value = currentAward.Id;
                    ws.Cells["B" + (printLine)].Value = currentAward.IsElite;
                    ws.Cells["C" + (printLine)].Value = currentAward.GoldCost;
                    ws.Cells["D" + (printLine)].Value = currentAward.Exp;
                }

                pck.Save();
            }

            if (File.Exists(newFile))
                return new FileNameOperationResponse(true, newFile);

            return new FileNameOperationResponse(false);
        }
        #endregion

        #region Utilities
        private static bool RunWithErrorHandling(Action action)
        {
            try
            {
                action.Invoke();
                return true;
            }
            catch (Exception ex)
            {
                WriteInColor("There was an error: " + ex.ToString(), ConsoleColor.Red);
                return false;
            }

        }

        private static void WriteInColor(string message, ConsoleColor color)
        {
            var defaultColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ForegroundColor = defaultColor;
        }
        #endregion


    }



}
